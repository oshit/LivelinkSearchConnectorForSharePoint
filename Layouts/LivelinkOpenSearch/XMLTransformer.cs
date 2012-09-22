using System;
using System.Globalization;
using System.IO;
using System.Web.UI;
using System.Xml.XPath;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    class XMLTransformer : Transformer {

        // Constructor. About the parameters, see the protected properties in the parent class.
        public XMLTransformer(string query, string searchUrl, string descriptorUrl) :
            base(query, searchUrl, descriptorUrl) {}

        // Transforms the Livelink XML Search results to the RSS 2.0 schema specialized
        // for OpenSearch 1.1 and writes it to the response output.
        public override void TransformResults(Stream results, HtmlTextWriter writer) {
            // If something goes severely wrong Livelink returns a HTML page with the error
            // details. In such case the following statement fails with an error that no
            // document can be recognoized in the input stream. It can happen if the HTTP
            // request was not autenticated - a login HTML page would be returned.
            var document = new XPathDocument(results);
            var navigator = document.CreateNavigator();
            // Livelink XML Search API returns errors in the XML output; not by the HTTP status.
            var error = navigator.SelectSingleNode("/Output/Error");
            if (error != null)
                throw new ApplicationException(error.GetSafeValue());
            var urlBase = SearchUrl.GetComponents(UriComponents.SchemeAndServer,
                UriFormat.Unescaped);
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            // OpenSearch schema enables paging of the search results. Yahoo Media schema
            // enables media content like image thumbnails. SharePoint Search schema enables
            // file type recognition.
            writer.Write("<rss version=\"2.0\"");
            writer.Write(" xmlns:os=\"http://a9.com/-/spec/opensearch/1.1/\"");
            writer.Write(" xmlns:m=\"http://search.yahoo.com/mrss/\"");
            writer.Write(" xmlns:ss=\"http://schemas.microsoft.com/SharePoint/Search/RSS\">");
            writer.Write("<channel>");
            // The title includes minimum information - the Livelink server host and query terms.
            writer.Write("<title>");
            writer.WriteEncodedText("Livelink Enterprise at ");
            writer.WriteEncodedText(SearchUrl.Host);
            writer.WriteEncodedText(": ");
            writer.WriteEncodedText(Query);
            writer.Write("</title>");
            writer.Write("<link rel=\"search\" type=\"application/opensearchdescription+xml\" ");
            writer.Write("href=\"");
            writer.WriteEncodedText(DescriptorUrl);
            writer.Write("\" title=\"");
            writer.WriteEncodedText("Livelink Enterprise at ");
            writer.WriteEncodedText(SearchUrl.Host);
            writer.Write("\" />");
            // The link should reproduce these query results if used interactively in the browser.
            writer.Write("<link>");
            writer.WriteEncodedText(SearchUrl.ToString());
            writer.Write("</link>");
            // The description contains the full original Livelink XML Search API URL.
            writer.Write("<description>");
            writer.WriteEncodedText("Search results of the query executed against the Enterprise Workspace of the Livelink server by the URL ");
            writer.WriteEncodedText(SearchUrl.ToString());
            writer.WriteEncodedText(".");
            writer.Write("</description>");
            // The name of this connector is used as the user agent for the Livelink requests
            // and also as the generator of these search results.
            writer.Write("<generator>");
            writer.WriteEncodedText(WebClient.UserAgent);
            writer.Write("</generator>");
            // The search results can be marked by the Livelink search icon pointing to the link
            // that should reproduce these query results interactively in the browser.
            writer.Write("<image>");
            writer.Write("<title>");
            writer.WriteEncodedText("Search results: ");
            writer.WriteEncodedText(Query);
            writer.Write("</title>");
            writer.Write("<url>");
            writer.WriteEncodedText(urlBase + "/img/icon-search.gif");
            writer.Write("</url>");
            writer.Write("<link>");
            writer.WriteEncodedText(SearchUrl.ToString());
            writer.Write("</link>");
            writer.Write("</image>");
            var info = navigator.SelectSingleNode("/Output/SearchResultsInformation");
            if (info != null) {
                // The following three elements are needed to enable paging in the search results.
                writer.Write("<os:totalResults>");
                writer.Write(info.SelectSingleNode("EstTotalResults").GetSafeValue());
                writer.Write("</os:totalResults>");
                writer.Write("<os:startIndex>");
                var start = info.SelectSingleNode("CurrentStartAt").GetSafeValue();
                int startAt = 0;
                if (!string.IsNullOrEmpty(start))
                    startAt = int.Parse(start, CultureInfo.InvariantCulture) - 1;
                writer.Write(start);
                writer.Write("</os:startIndex>");
                writer.Write("<os:itemsPerPage>");
                var count = info.SelectSingleNode("NumberResultsThisPage").GetSafeValue();
                writer.Write(count);
                writer.Write("</os:itemsPerPage>");
                // Extra query information in the OpenSearch schema.
                writer.Write("<os:Query role=\"request\" title=\"");
                writer.WriteEncodedText("Livelink Search");
                writer.Write("\" searchTerms=\"");
                writer.WriteEncodedText(Query);
                writer.Write("\" startPage=\"");
                var pageSize = int.Parse(count, CultureInfo.InvariantCulture);
                var pageIndex = int.Parse(start, CultureInfo.InvariantCulture);
                if (pageSize > 0)
                    pageIndex = pageIndex / pageSize + 1;
                writer.Write(pageIndex);
                writer.Write("\" />");
            }
            // The search hits follow not encapsulated in an XML element.
            var hits = navigator.Select("/Output/SearchResults/SearchResult");
            foreach (XPathNavigator hit in hits) {
                // Although no hits are returned there is one element SearchResult coming
                // containing the text "Sorry, no results were found".
                var name = hit.SelectSingleNode("OTName");
                if (name == null)
                    break;
                writer.Write("<item>");
                // The title of the hit is made of the Livelink object name which is usually
                // the file name for documents.
                writer.Write("<title>");
                // OTName element can have language-dependent sub-elements; take just the body.
                var title = name.SelectSingleNode("text()").GetSafeValue();
                writer.WriteEncodedText(title);
                writer.Write("</title>");
                // The link points to the Livelink object's viewing URL which is usually
                // the overview page with basic properties and a download link.
                writer.Write("<link>");
                // OTName should always be present but just be on the safe side here.
                if (name != null) {
                    var url = urlBase + name.GetAttribute("ViewURL", "");
                    writer.WriteEncodedText(url);
                }
                writer.Write("</link>");
                // The description contains a document fragment where the search terms were
                // found. It is available for documents only; objects without content that
                // were returned because thir name or other meta-data matched have this
                // value empty.
                writer.Write("<description>");
                var summary = hit.SelectSingleNode("OTSummary").GetSafeValue();
                if (MaxSummaryLength > 0 && summary.Length > MaxSummaryLength)
                    summary = summary.Substring(0, MaxSummaryLength) + " ...";
                summary = summary.Replace("<HH>", "<b>").Replace("</HH>", "</b>");
                writer.WriteEncodedText(summary);
                writer.Write("</description>");
                // Publishing date should be the last modificattion time but Livelink returns
                // just the creation date from the default search template.
                var date = hit.SelectSingleNode("OTObjectDate");
                if (date != null) {
                    writer.Write("<pubDate>");
                    writer.WriteEncodedText(date.GetSafeValue());
                    writer.Write("</pubDate>");
                }
                // Unfortunately, Livelink returns only the numeric user identifier
                // of the document's creator here. Resolving the actual user name
                // by additional HTTP requests woudl be too bad for the performance.
                var author = hit.SelectSingleNode("OTCreatedBy");
                if (author != null) {
                    writer.Write("<author>");
                    writer.WriteEncodedText(author.GetSafeValue());
                    writer.Write("</author>");
                }
                // SharePoint sadly does not recognize the thumbnail and enclosure elements
                // to display icon and basic document information. Windows search does.
                var type = hit.SelectSingleNode("OTMIMEType");
                // There is no icon without the MIME type element which is returned for all object
                // types, surprisingly. The text vaue is empty if there is no actual  MIME type.
                if (type != null) {
                    writer.Write("<m:thumbnail url=\"");
                    var url = urlBase + type.GetAttribute("IconURL", "");
                    writer.WriteEncodedText(url);
                    writer.Write("\" />");
                }
                // The document size should be written in bytes to the output.
                var size = hit.SelectSingleNode("OTObjectSize");
                var length = 0;
                if (size != null) {
                    var number = int.Parse(size.GetSafeValue(), CultureInfo.InvariantCulture);
                    var factor = 1;
                    var suffix = size.GetAttribute("Suffix", "");
                    if (suffix != null)
                        switch (suffix.Trim().ToUpperInvariant()) {
                            case "KB":
                                factor = 1024;
                                break;
                            case "MB":
                                factor = 1024 * 1024;
                                break;
                            case "GB":
                                factor = 1024 * 1024 * 1024;
                                break;
                        }
                    length = number * factor;
                }
                // Both SharePoint and Livelink do not allow documents with zero length.
                // Size and file extension properties should be written only for documents.
                if (length > 0) {
                    writer.Write("<ss:size>");
                    writer.Write(length);
                    writer.Write("</ss:size>");
                    var dot = title.LastIndexOf('.');
                    if (dot >= 0) {
                        writer.Write("<ss:dotfileextension>");
                        writer.Write(title.Substring(dot).ToUpper());
                        writer.Write("</ss:dotfileextension>");
                    }
                }
                // Clients that can use extra content information can show a direct download
                // link to open the document without first visiting the overview page.
                if (name != null && type != null && !string.IsNullOrWhiteSpace(type.Value)) {
                    writer.Write("<enclosure url=\"");
                    var url = urlBase + name.GetAttribute("DownloadURL", "");
                    writer.WriteEncodedText(url);
                    writer.Write("\" type=\"");
                    writer.WriteEncodedText(type.GetSafeValue());
                    if (length > 0) {
                        writer.Write("\" length=\"");
                        writer.Write(length);
                    }
                    writer.Write("\" />");
                }
                writer.Write("</item>");
            }
            writer.Write("</channel>");
            writer.Write("</rss>");
        }
    }
}
