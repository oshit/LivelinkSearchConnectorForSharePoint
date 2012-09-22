using System;
using System.Globalization;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Xml.XPath;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    class HTMLTransformer : Transformer {

        // The URL to get the previous page of search results or null if there is none.
        string PreviousUrl { get; set; }

        // The URL to get the next page of search results or null if there is none.
        string NextUrl { get; set; }

        // Constructor. About the parameters, see the protected properties in the parent class
        // and also the properties in this class declared above.
        public HTMLTransformer(string query, string searchUrl, string previousUrl,
                string nextUrl, string descriptorUrl) : base(query, searchUrl, descriptorUrl) {
            PreviousUrl = previousUrl;
            NextUrl = nextUrl;
        }

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
            writer.WriteFullBeginTag("html");
            writer.WriteFullBeginTag("head");
            writer.WriteFullBeginTag("title");
            writer.WriteEncodedText("Search Results at ");
            writer.WriteEncodedText(SearchUrl.Host);
            writer.WriteEndTag("title");
            writer.WriteFullBeginTag("style");
            writer.WriteEncodedText(@"
body { font-family: Arial, sans-serif }
.summary { margin-top: 0.5em }
.data { margin-top: 0.3em; color: green }
.summary, .note, .data { font-size: 85% }
td { padding-right: 0.3em }
dl { margin-top: 1em }
dt { margin-top: 1em }
dd { margin-left: 0 }
");
            writer.WriteEncodedText(SearchUrl.Host);
            writer.WriteEndTag("style");
            writer.WriteBeginTag("link");
            writer.WriteAttribute("type", "application/opensearchdescription+xml");
            writer.WriteAttribute("rel", "search");
            writer.WriteAttribute("href", DescriptorUrl, true);
            writer.WriteAttribute("title", "Search Results at " + SearchUrl.Host, true);
            writer.Write(HtmlTextWriter.TagRightChar);
            string startAt = null;
            string actualPageSize = null;
            string totalCount = null;
            var info = navigator.SelectSingleNode("/Output/SearchResultsInformation");
            if (info != null) {
                startAt = info.SelectSingleNode("CurrentStartAt").GetSafeValue();
                actualPageSize = info.SelectSingleNode("NumberResultsThisPage").GetSafeValue();
                totalCount = info.SelectSingleNode("EstTotalResults").GetSafeValue();
                writer.WriteBeginTag("meta");
                writer.WriteAttribute("name", "totalResults");
                writer.WriteAttribute("content", totalCount, true);
                writer.Write(HtmlTextWriter.TagRightChar);
                writer.WriteBeginTag("meta");
                writer.WriteAttribute("name", "startIndex");
                writer.WriteAttribute("content", startAt, true);
                writer.Write(HtmlTextWriter.TagRightChar);
                writer.WriteBeginTag("meta");
                writer.WriteAttribute("name", "itemsPerPage");
                writer.WriteAttribute("content", actualPageSize, true);
                writer.Write(HtmlTextWriter.TagRightChar);
            }
            writer.WriteEndTag("head");
            writer.WriteFullBeginTag("body");
            // Page title: icon and caption.
            writer.WriteFullBeginTag("table");
            writer.WriteFullBeginTag("tr");
            writer.WriteFullBeginTag("td");
            writer.WriteBeginTag("img");
            writer.WriteAttribute("src", urlBase + "/img/icon_search.gif", true);
            writer.WriteAttribute("alt", "Search results: " + Query, true);
            writer.Write(HtmlTextWriter.TagRightChar);
            writer.WriteEndTag("td");
            writer.WriteFullBeginTag("td");
            writer.WriteFullBeginTag("h2");
            writer.WriteEncodedText("Results of the search for ");
            writer.WriteEncodedText(Query);
            writer.WriteEndTag("h2");
            writer.WriteEndTag("td");
            writer.WriteEndTag("tr");
            writer.WriteEndTag("table");
            // Note about the source of the search results.
            writer.WriteBeginTag("div");
            writer.WriteAttribute("class", "note", true);
            writer.Write(HtmlTextWriter.TagRightChar);
            writer.WriteEncodedText("Search results of the query executed against the Enterprise Workspace of the Livelink server by the URL ");
            writer.WriteBeginTag("a");
            writer.WriteAttribute("href", SearchUrl.ToString(), true);
            writer.Write(HtmlTextWriter.TagRightChar);
            writer.WriteEncodedText(SearchUrl.ToString());
            writer.WriteEndTag("a");
            writer.WriteEncodedText(". The OSDX file for this search source can be ");
            writer.WriteBeginTag("a");
            writer.WriteAttribute("href", DescriptorUrl, true);
            writer.Write(HtmlTextWriter.TagRightChar);
            writer.WriteEncodedText("downloaded for your OpenSearch client");
            writer.WriteEndTag("a");
            writer.WriteEncodedText(" or ");
            writer.WriteBeginTag("a");
            writer.WriteAttribute("href", "#");
            writer.WriteAttribute("onclick", string.Format(
                "window.external.AddSearchProvider('{0}')",
                HttpUtility.HtmlEncode(DescriptorUrl)));
            writer.Write(HtmlTextWriter.TagRightChar);
            writer.WriteEncodedText("installed in the browser");
            writer.WriteEndTag("a");
            writer.WriteEncodedText(".");
            writer.WriteEndTag("div");
            var hits = navigator.Select("/Output/SearchResults/SearchResult");
            if (hits.Count > 0) {
                // Search hits as definition list.
                writer.WriteFullBeginTag("dl");
                foreach (XPathNavigator hit in hits) {
                    // Although no hits are returned there is one element SearchResult coming
                    // containing the text "Sorry, no results were found".
                    var name = hit.SelectSingleNode("OTName");
                    if (name == null) {
                        writer.WriteFullBeginTag("dt");
                        writer.WriteEncodedText(hit.GetSafeValue());
                        writer.WriteEndTag("dt");
                        break;
                    }
                    writer.WriteFullBeginTag("dt");
                    writer.WriteFullBeginTag("table");
                    writer.WriteFullBeginTag("tr");
                    writer.WriteFullBeginTag("td");
                    var type = hit.SelectSingleNode("OTMIMEType");
                    // There is no icon without the MIME type element which is returned for all object
                    // types, surprisingly. The text vaue is empty if there is no actual  MIME type.
                    if (type != null) {
                        writer.WriteBeginTag("img");
                        var url = urlBase + type.GetAttribute("IconURL", "");
                        writer.WriteAttribute("src", url, true);
                        writer.WriteAttribute("alt", type.GetSafeValue(), true);
                        writer.Write(HtmlTextWriter.TagRightChar);
                        writer.WriteEndTag("td");
                        writer.WriteFullBeginTag("td");
                    }
                    // OTName should always be present but just be on the safe side here.
                    if (name != null) {
                        writer.WriteBeginTag("a");
                        // The link points to the Livelink object's viewing URL which is usually
                        // the overview page with basic properties and a download link.
                        var url = urlBase + name.GetAttribute("ViewURL", "");
                        writer.WriteAttribute("href", url, true);
                        writer.Write(HtmlTextWriter.TagRightChar);
                        // The title of the hit is made of the Livelink object name which is usually
                        // the file name for documents. OTName element can have language-dependent
                        // sub-elements; take just the body.
                        var title = name.SelectSingleNode("text()").GetSafeValue();
                        writer.WriteEncodedText(title);
                        writer.WriteEndTag("a");
                    }
                    writer.WriteEndTag("td");
                    writer.WriteEndTag("tr");
                    writer.WriteEndTag("table");
                    writer.WriteEndTag("dt");
                    writer.WriteFullBeginTag("dd");
                    writer.WriteBeginTag("div");
                    writer.WriteAttribute("class", "summary", true);
                    writer.Write(HtmlTextWriter.TagRightChar);
                    // The description contains a document fragment where the search terms were
                    // found. It is available for documents only; objects without content that
                    // were returned because thir name or other meta-data matched have this
                    // value empty.
                    var summary = hit.SelectSingleNode("OTSummary").GetSafeValue();
                    if (MaxSummaryLength > 0 && summary.Length > MaxSummaryLength)
                        summary = summary.Substring(0, MaxSummaryLength) + " ...";
                    summary = summary.Replace("<HH>", "<b>").Replace("</HH>", "</b>");
                    writer.Write(summary);
                    writer.WriteEndTag("div");
                    writer.WriteBeginTag("div");
                    writer.WriteAttribute("class", "data", true);
                    writer.Write(HtmlTextWriter.TagRightChar);
                    // Only documents have their byte size.
                    var size = hit.SelectSingleNode("OTObjectSize");
                    if (size != null) {
                        writer.WriteEncodedText("Size: ");
                        writer.WriteEncodedText(size.GetSafeValue());
                        writer.WriteEncodedText(" ");
                        var suffix = size.GetAttribute("Suffix", "");
                        if (string.IsNullOrEmpty(suffix))
                            suffix = "B";
                        writer.WriteEncodedText(suffix);
                        writer.WriteEncodedText(". ");
                    }
                    // The description contains a document fragment where the search terms were
                    // found. It is available for documents only; objects without content that
                    // were returned because thir name or other meta-data matched have this
                    // value empty.
                    writer.WriteEncodedText("Created by ");
                    // Unfortunately, Livelink returns only the numeric user identifier
                    // of the document's creator here. Resolving the actual user name
                    // by additional HTTP requests woudl be too bad for the performance.
                    var author = hit.SelectSingleNode("OTCreatedBy");
                    if (author != null)
                        writer.WriteEncodedText(author.GetSafeValue());
                    writer.WriteEncodedText(" at ");
                    // Publishing date should be the last modificattion time but Livelink returns
                    // just the creation date from the default search template.
                    var date = hit.SelectSingleNode("OTObjectDate");
                    if (date != null)
                        writer.WriteEncodedText(date.GetSafeValue());
                    writer.WriteEncodedText(". ");
                    writer.WriteEncodedText("Location: ");
                    var location = hit.SelectSingleNode("OTLocation");
                    if (location != null) {
                        writer.WriteBeginTag("a");
                        var url = urlBase + location.GetAttribute("URL", "");
                        writer.WriteAttribute("href", url, true);
                        writer.Write(HtmlTextWriter.TagRightChar);
                        writer.WriteEncodedText(location.GetAttribute("Name", ""));
                        writer.WriteEndTag("a");
                    }
                    writer.WriteEncodedText(".");
                    writer.WriteEndTag("div");
                    writer.WriteEndTag("dd");
                }
                writer.WriteEndTag("dl");
                if (info != null) {
                    // Search results paging control.
                    writer.WriteFullBeginTag("center");
                    writer.WriteFullBeginTag("table");
                    writer.WriteFullBeginTag("tr");
                    if (PreviousUrl != null) {
                        writer.WriteFullBeginTag("td");
                        writer.WriteBeginTag("a");
                        writer.WriteAttribute("href", PreviousUrl, true);
                        writer.Write(HtmlTextWriter.TagRightChar);
                        writer.WriteBeginTag("img");
                        writer.WriteAttribute("src", urlBase + "/img/page_previous16.gif", true);
                        writer.WriteAttribute("alt", "Previous Page", true);
                        writer.Write(HtmlTextWriter.TagRightChar);
                        writer.WriteEndTag("a");
                        writer.WriteEndTag("td");
                    }
                    writer.WriteFullBeginTag("td");
                    writer.WriteEncodedText("Items ");
                    writer.Write(startAt);
                    writer.Write(" - ");
                    var endAt = int.Parse(startAt, CultureInfo.InvariantCulture) +
                        int.Parse(actualPageSize, CultureInfo.InvariantCulture) - 1;
                    writer.Write(endAt);
                    writer.Write(" of ");
                    writer.Write(totalCount);
                    writer.WriteEndTag("td");
                    if (NextUrl != null) {
                        var pageSize = SearchHelper.GetRequestedItemCount(SearchUrl.Query);
                        if (!string.IsNullOrEmpty(pageSize) &&
                            !string.IsNullOrEmpty(actualPageSize)) {
                            var count = int.Parse(pageSize, CultureInfo.InvariantCulture);
                            var actualCount = int.Parse(actualPageSize,
                                CultureInfo.InvariantCulture);
                            if (count == actualCount) {
                                writer.WriteFullBeginTag("td");
                                writer.WriteBeginTag("a");
                                writer.WriteAttribute("href", NextUrl, true);
                                writer.Write(HtmlTextWriter.TagRightChar);
                                writer.WriteBeginTag("img");
                                writer.WriteAttribute("src", urlBase + "/img/page_next16.gif", true);
                                writer.WriteAttribute("alt", "Next Page", true);
                                writer.Write(HtmlTextWriter.TagRightChar);
                                writer.WriteEndTag("a");
                                writer.WriteEndTag("td");
                            }
                        }
                    }
                    writer.WriteEndTag("tr");
                    writer.WriteEndTag("table");
                    writer.WriteEndTag("center");
                }
            } else {
                writer.WriteFullBeginTag("p");
                writer.WriteEncodedText("No items were found.");
                writer.WriteEndTag("p");
            }
            writer.WriteEndTag("body");
            writer.WriteEndTag("html");
        }
    }
}
