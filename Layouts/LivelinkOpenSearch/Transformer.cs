using System;
using System.IO;
using System.Web.UI;
using System.Xml.XPath;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    abstract class Transformer {

        // Can limit the maximum length of the textual summary that is displayed below a search
        // hit to give a hint what is the document about. It usually contains highlighted
        // search terms. I saw Livelink returning really long passages (more than 1300 characters)
        // while I have not seen SharePoint returning more than 185 characters. It can make the
        // result list long and difficult to browse. You can trim the length of the displayed
        // summary by this parameter to never exceed the specified length.
        public int MaxSummaryLength { get; set; }

        // Search terms in the syntax recognized by the Livelink XML Search API. It can be
        // a list of words that you want to be found in the returned search results for,
        // for example.
        protected string Query { get; private set; }

        // The complete URL (including the query part) that can be navigated to in the browser
        // to get the Livelink search results (in HTML) that would be processed by this class.
        protected Uri SearchUrl { get; private set; }

        // The URL of the OpenSearch descriptor (OSDX) file.
        protected string DescriptorUrl { get; private set; }

        // Constructor. About the parameters, see the properties above.
        protected Transformer(string query, string searchUrl, string descriptorUrl) {
            if (query == null)
                throw new ArgumentNullException("query");
            if (searchUrl == null)
                throw new ArgumentNullException("searchUrl");
            if (descriptorUrl == null)
                throw new ArgumentNullException("descriptorUrl");
            Query = query;
            SearchUrl = new Uri(searchUrl);
            DescriptorUrl = descriptorUrl;
        }

        // Transforms the Livelink XML Search results to the format defined by the descendant
        // class and writes it to the response output.
        public abstract void TransformResults(Stream results, HtmlTextWriter writer);
    }

    // Extends the class XPathNavigator with extra methods, for example:
    //
    //   // Always receives a string; regardless the XPath existence.
    //   var value = navigator.SelectSingleNode("Value").GetSafeValue();
    static class XPathNavigatorExtension {

        // Returns the textual value of the element the navigator points to. If the navigator
        // itself or the value of the element are null an empty string is returned. Whitespace
        // at the the beginning and at the end of the returned value is always trimmed.
        public static string GetSafeValue(this XPathNavigator navigator) {
            if (navigator != null) {
                var value = navigator.Value;
                if (value != null)
                    return value.Trim();
            }
            return "";
        }
    }
}
