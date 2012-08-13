using System;
using System.Text;
using System.Web;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // Builds URL queries for Livelink XML search API requests. The query can be
    // directly used in a HTTP request to Livelink that is expected to return XML
    // results. Typical usage:
    //
    //   var helper = new SearchHelper();
    //   var query = helper.GetUrlQuery("XML Search", 1, 10);
    class SearchHelper {

        // Additional query string to be appended to the end of the search query
        // string. The string should not start with ampersand and will be appended as-is.
        public string AdditionalParameters { get; set; }

        // User to impersonate the query for. If impersonated queries are used the HTTP
        // request must be authenticated as a Livelink system administrator.
        string User { get; set; }

        // Constructor. No explicit user is needed for the SSO scenario.
        public SearchHelper() {}

        // Constructor. About the parameter user, see the property User above.
        public SearchHelper(string user) : this() {
            if (user == null)
                throw new ArgumentNullException("user");
            User = user;
        }

        // Returns the complete URL query that can be used to perform a search by the Livelink
        // XML Search API. The query should contain the search terms recognized by Livelink.
        public string GetUrlQuery(string query, int startIndex = 0, int count = 0) {
            if (query == null)
                throw new ArgumentNullException("query");
            var result = new StringBuilder("func=search&outputformat=xml&where1=");
            result.Append(HttpUtility.UrlEncode(query));
            if (User != null)
                result.Append("&userLogin=").Append(HttpUtility.UrlEncode(User));
            if (startIndex > 0)
                result.Append("&startat=").Append(startIndex);
            if (count > 0)
                result.Append("&gofor=").Append(count);
            if (!string.IsNullOrEmpty(AdditionalParameters))
                result.Append("&").Append(AdditionalParameters);
            return result.ToString();
        }

        // Takes a URL query of a Livelink search request and makes sure that the returned
        // value will produces a HTML output and will contain no user impersonation.
        public static string ConvertToBrowserUsage(string query) {
            if (query == null)
                throw new ArgumentNullException("query");
            var start = query.IndexOf("&userLogin=");
            if (start >= 0) {
                var end = query.IndexOf('&', start + 11);
                if (end >= 0)
                    query = query.Remove(start, end - start);
                else
                    query.Remove(start);
            }
            return query.Replace("&outputformat=xml", "");
        }
    }
}
