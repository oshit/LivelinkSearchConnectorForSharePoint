using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Web.UI;
using System.Xml.XPath;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    class HTMLTransformer : Transformer {

        // The URL to get the previous page of search results or null if there is none.
        string PreviousUrl { get; set; }

        // The URL to get the next page of search results or null if there is none.
        string NextUrl { get; set; }

        // Constructor. About the parameters, see the protected properties in the parent class
        // and also the properties in this class declared above.
        public HTMLTransformer(string query, string searchUrl, string previousUrl, string nextUrl)
                : base(query, searchUrl) {
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
.summary, .data { margin-top: 0.3em }
.data { color: green }
.summary, .note, .data { font-size: 85% }
td { padding-right: 0.2em }
dl { margin-top: 1em }
dt { margin-top: 1em; margin-bottom: 0.3em }
dd { margin-left: 0 }
");
            writer.WriteEncodedText(SearchUrl.Host);
            writer.WriteEndTag("style");
            writer.WriteEndTag("head");
            writer.WriteFullBeginTag("body");
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
            writer.WriteBeginTag("div");
            writer.WriteAttribute("class", "note", true);
            writer.Write(HtmlTextWriter.TagRightChar);
            writer.WriteEncodedText("Search results of the query executed against the Enterprise Workspace of the Livelink server by the URL ");
            writer.WriteBeginTag("a");
            writer.WriteAttribute("href", SearchUrl.ToString(), true);
            writer.Write(HtmlTextWriter.TagRightChar);
            writer.WriteEncodedText(SearchUrl.ToString());
            writer.WriteEndTag("a");
            writer.WriteEncodedText(".");
            writer.WriteEndTag("div");
            var hits = navigator.Select("/Output/SearchResults/SearchResult");
            if (hits.Count > 0) {
                writer.WriteFullBeginTag("dl");
                foreach (XPathNavigator hit in hits) {
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
                    var name = hit.SelectSingleNode("OTName");
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
                    writer.WriteEncodedText(summary);
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
                var info = navigator.SelectSingleNode("/Output/SearchResultsInformation");
                if (info != null) {
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
                    var startAt = info.SelectSingleNode("CurrentStartAt").GetSafeValue();
                    writer.Write(startAt);
                    writer.Write(" - ");
                    var actualPageSize = info.SelectSingleNode(
                        "NumberResultsThisPage").GetSafeValue();
                    var endAt = int.Parse(startAt, CultureInfo.InvariantCulture) +
                        int.Parse(actualPageSize, CultureInfo.InvariantCulture) - 1;
                    writer.Write(endAt);
                    writer.Write(" of ");
                    writer.Write(info.SelectSingleNode("EstTotalResults").GetSafeValue());
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
