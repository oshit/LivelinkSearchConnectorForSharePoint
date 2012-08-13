using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Web.UI;
using System.Xml.XPath;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // Performs Livelink XML Search queries and returns the result as a RSS 2.0 content.
    // If the operation fails the page returns the HTTP status code 500 with the error
    // message in the status description.
    //
    // If Kerberos is used SSO can be turned on by UseSSO=true; the current authenticated
    // SharePoint user context will be passed to the Livelink HTTP request. Otherwise
    // the Livelink search request will be performed as a Livelink system administrator
    // with credentials provided by the SharePoint Secure Store Service. The call will be
    // impersonated to a Livelink user mapped from the current SharePoint user to trim the
    // search results according to the current user's permissions.
    //
    // For example, the following parameters will make the URL below:
    //   Query             - XML Search
    //   LivelinkUrl       - http://myserver/livelink/llisapi.dll
    //   TargetAppID       - myserver
    //   LoginPattern      - {user:lc}
    //   StartIndex        - 1
    //   Count             - 10
    //   ExtraParams       - lookfor1=allwords
    //   IgnoreSSLWarnings - true
    //
    //   http://sparepoint/_layouts/LivelinkOpenSearch/ExecuteQuery?query=XML+Search&
    //   livelinkUrl=http%3A%2F%2Fmyserver%2Flivelink%2Fllisapi.dll&targetAppID=myserver&
    //   loginPattern=%7Buser:lc%7D&startIndex=1&count=10&extraParams=lookfor1%3Dallwords&
    //   ignoreSSLWarnings=true
    //
    // Notice that some parameters are URL parts and should be URL encoded so that this page
    // gets them correctly when it is called. The parameter LoginPattern is URL encoded too;
    // some expressions in braces could be mistakenly resolved by the search federator before
    // it calls the URL.
    public partial class ExecuteQueryPage : LayoutsPageBase {

        // Search terms in the syntax recognized by the Livelink XML Search API. It can be
        // a list of words that you want to be found in the returned search results for,
        // for example. This parameter is mandatory.
        string Query { get; set; }

        // The complete URL of the Livelink CGI executable or the Livelink module for the web
        // server you use; http://myserver/livelink/llisapi.dll, for example.
        // This parameter is mandatory.
        string LivelinkUrl { get; set; }

        // The identifier of a target application in the SharePoint Secure Store that contains
        // credentials of a Livelink system administrator. Only the value types UserName
        // and Password are read from there and they must not be empty.
        //
        // If a search query needs to be impersonated for another user a system administrator
        // must execute it. If UseSSO is not provided or is false this parameter is mandatory
        // otherwise it is ignored and may be omitted.
        string TargetAppID { get; set; }

        // The text pattern to generate a Livelink login name from a SharePoint user. Users are
        // usually synchronized from SharePoint (Windows AD) to Livelink and their names are
        // transformed by using a predictable pattern. For example, the following pattern will
        // use the SharePoint user name without the Windows domain converted to lower-case
        // and append to it a constant e-mail-like domain: "{user:lc}@myserver.com".
        //
        // The pattern describing the user name transformation is a combination of parameter
        // placeholders and other text. The parameter placeholders are enclosed in braces
        // and are replaced with values from the SharePoint user information:
        //
        //   {login}  - complete login name; with the domain if provided.
        //   {user}   - only user name; without the domain.
        //   {domain} - only the domain name; empty if no domain is provided.
        //
        // Unrecognized parameter placeholders (expressions in braces) are left intact.
        // The parameter placeholders can be used together with modifiers that transform the text
        // value further. They start with colon and are written between the parameter placeholder
        // name and the right closing brace like {domain:uc}, for example:
        //
        //   :lc - transform the value to lower-case
        //   :uc - transform the value to upper-case
        //
        // If a search query needs to be impersonated for another user its Livelink login name
        // must be used. If UseSSO is not provided or is false this parameter is mandatory
        // otherwise it is ignored and may be omitted.
        string LoginPattern { get; set; }

        // If you use SSO in the entire environment covering Windows, SharePoint and Livelink
        // that allows you working in the browser with any system without an explicit login
        // you can set this parameter to true. You will need neither configuring the Livelink
        // system administrator credentials in a SharePoint Secure Store target application
        // nor the login name pattern to infer the Livelink user name from the authenticated
        // SharePoint user. The default value is false.
        bool UseSSO { get; set; }

        // If the Livelink LivelinkUrl includes the HTTPS protocol and the SSL certificate
        // of the Livelink server is invalid (expired, for example) the search request
        // would fail. If you are sure and want to connect to the Livelink server
        // in spite of it set this parameter to true. The default value is false.
        bool IgnoreSSLWarnings { get; set; }

        // Start index of the first search result to return. If not provided the number 1 is
        // assumed; to start with the very first search hit available.
        int StartIndex { get; set; }

        // The maximum count of returned search results. If not provided the actual count
        // will be decided by the search engine. Usually not all available hits are returned;
        // the maximum count can be limited to 1000 from peformance reasons, for example.
        int Count { get; set; }

        // Additional URL query string to append to the Livelink XML Search API URL.
        // The string should not start with ampersand and will be appended as-is.
        // For example, the value "lookfor1=allwords" will require returning only those
        // results where all entered search terms can be found.
        string ExtraParams { get; set; }

        // The encoding of the input search terms. It is used to check that only either ASCII
        // or UTF-8 URL parameters were sent. Other encodings are not supported now.
        string InputEncoding { get; set; }

        // The expected encoding of the search results. It is used to check that only either
        // ASCII or UTF-8 output was requested. Other encodings are not supported now.
        string OutputEncoding { get; set; }

        // The preferred language for the search results. It is currently unused.
        string Language { get; set; }

        // Reads and checks the provided URL parameters, performs the Livelink XML Search
        // query, transforms its results to RSS 2.0 and writes the content to the response
        // output as text/xml. If anything fails the HTTP status code 500 will be returned
        // in the response with the error message in the status description.
        protected override void Render(HtmlTextWriter writer) {
            try {
                ParseUrl();
                if (IgnoreSSLWarnings)
                    ServicePointManager.ServerCertificateValidationCallback =
                        (sender, cert, chain, sslError) => true;
                var query = GetQuery();
                using (var results = GetResults(query)) {
                    Response.ContentType = "text/xml";
                    TransformResults(writer, query, results);
                }
            } catch (Exception exception) {
                Response.StatusCode = 500;
                Response.StatusDescription = exception.Message;
            }
        }

        // Sets values of properties in this class from the provided URL parameters
        // or from the assumed parameter defaults.
        void ParseUrl() {
            Query = Request.QueryString["query"];
            if (string.IsNullOrEmpty(Query))
                throw new ApplicationException("Search query was empty.");
            LivelinkUrl = Request.QueryString["livelinkUrl"];
            if (string.IsNullOrEmpty(LivelinkUrl))
                throw new ApplicationException("Livelink CGI URL was empty.");
            UseSSO = "true".Equals(Request.QueryString["useSSO"],
                StringComparison.InvariantCultureIgnoreCase);
            // If SSO is not enabled the user impersonation parameters are ignored.
            if (!UseSSO) {
                TargetAppID = Request.QueryString["targetAppID"];
                if (string.IsNullOrEmpty(TargetAppID))
                    throw new ApplicationException("Target application ID was empty.");
                LoginPattern = Request.QueryString["loginPattern"];
                if (string.IsNullOrEmpty(LoginPattern))
                    throw new ApplicationException("User login pattern was empty.");
            }
            IgnoreSSLWarnings = "true".Equals(Request.QueryString["IgnoreSSLWarnings"],
                StringComparison.InvariantCultureIgnoreCase);
            var startIndex = Request.QueryString["startIndex"];
            if (!string.IsNullOrEmpty(startIndex))
                StartIndex = int.Parse(startIndex, CultureInfo.InvariantCulture) + 1;
            var count = Request.QueryString["count"];
            if (!string.IsNullOrEmpty(count))
                Count = int.Parse(count, CultureInfo.InvariantCulture);
            ExtraParams = Request.QueryString["extraParams"];
            InputEncoding = Request.QueryString["inputEncoding"];
            if (!string.IsNullOrEmpty(InputEncoding) && !(
                "ASCII".Equals(InputEncoding, StringComparison.InvariantCultureIgnoreCase) ||
                "UTF-8".Equals(InputEncoding, StringComparison.InvariantCultureIgnoreCase)))
                throw new ApplicationException("The input encosing was neither ASCII nor UTF-8.");
            OutputEncoding = Request.QueryString["outputEncoding"];
            if (!string.IsNullOrEmpty(OutputEncoding) && !(
                "ASCII".Equals(OutputEncoding, StringComparison.InvariantCultureIgnoreCase) ||
                "UTF-8".Equals(OutputEncoding, StringComparison.InvariantCultureIgnoreCase)))
                throw new ApplicationException("The output encosing was neither ASCII nor UTF-8.");
            Language = Request.QueryString["language"];
        }

        // Computes the URL query for performing a search by the Livelink XML Search API.
        string GetQuery() {
            SearchHelper helper;
            if (UseSSO) {
                helper = new SearchHelper();
            } else {
                // If SSO is not enabled the search will be impersonated for the current user.
                var mapper = new LoginMapper(LoginPattern);
                var user = mapper.GetLoginName(SPContext.Current.Web.CurrentUser);
                helper = new SearchHelper(user);
            }
            if (!string.IsNullOrEmpty(ExtraParams))
                helper.AdditionalParameters = ExtraParams;
            return helper.GetUrlQuery(Query, StartIndex, Count);
        }

        // Performs a HTTP request using the specified URL query that is supposed to be a valid
        // Livelink XML Search API call returning XML content with the search results.
        // The returned stream must be disposed when not needed anymore.
        Stream GetResults(string query) {
            var client = new WebClient(LivelinkUrl) { UseSSO = UseSSO };
            // If SSO is not enabled another HTTP request is made before the searching one
            // to authenticate the operation. A search impersonated for (any) current user
            // can be performed only by a Livelink system administrator. The URL query passed
            // to this method is expected to contain the user login to trim the results for.
            if (!UseSSO) {
                // The SharePoint Secure Store Service will be accessed impersonated
                // to the privileged user running this code: the web application pool user.
                var store = new CredentialStore(SPContext.Current.Site, true);
                using (var credentials = store.GetCredentials(TargetAppID))
                    client.Authenticate(credentials);
            }
            var request = client.CreateRequest(query);
            return request.GetResponseContent();
        }

        // Transforms the Livelink XML Search results to the RSS 2.0 schema specialized
        // for OpenSearch 1.1 and writes it to the response output.
        void TransformResults(HtmlTextWriter writer, string query, Stream results) {
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
            var urlBase = new Uri(LivelinkUrl).GetComponents(UriComponents.SchemeAndServer,
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
            writer.WriteEncodedText(new Uri(LivelinkUrl).Host);
            writer.WriteEncodedText(": ");
            writer.WriteEncodedText(Query);
            writer.Write("</title>");
            // The link should reproduce these query results if used interactively in the browser.
            writer.Write("<link>");
            query = SearchHelper.ConvertToBrowserUsage(query);
            writer.WriteEncodedText(LivelinkUrl + "?" + query);
            writer.Write("</link>");
            // The description contains the full original Livelink XML Search API URL.
            writer.Write("<description>");
            writer.WriteEncodedText("Search results of the query executed against the Enterprise Workspace of the Livelink server by the URL ");
            writer.WriteEncodedText(LivelinkUrl + "?" + query);
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
            writer.WriteEncodedText(LivelinkUrl + "?" + query);
            writer.Write("</link>");
            writer.Write("</image>");
            var info = navigator.SelectSingleNode("/Output/SearchResultsInformation");
            if (info != null) {
                // The following three elements are needed to enable paging in the search results.
                writer.Write("<os:totalResults>");
                writer.Write(info.SelectSingleNode("EstTotalResults").GetSafeValue());
                writer.Write("</os:totalResults>");
                writer.Write("<os:startIndex>");
                writer.Write(info.SelectSingleNode("CurrentStartAt").GetSafeValue());
                writer.Write("</os:startIndex>");
                writer.Write("<os:itemsPerPage>");
                writer.Write(info.SelectSingleNode("NumberResultsThisPage").GetSafeValue());
                writer.Write("</os:itemsPerPage>");
                // Extra query information in the OpenSearch schema.
                writer.Write("<os:Query role=\"request\" title=\"");
                writer.WriteEncodedText("Livelink Search");
                writer.Write("\" searchTerms=\"");
                writer.WriteEncodedText(Query);
                writer.Write("\" startPage=\"");
                writer.Write(StartIndex > 0 ? StartIndex : 1);
                writer.Write("\" />");
            }
            // The search hits follow not encapsulated in an XML element.
            var hits = navigator.Select("/Output/SearchResults/SearchResult");
            foreach (XPathNavigator hit in hits) {
                writer.Write("<item>");
                // The title of the hit is made of the Livelink object name which is usually
                // the file name for documents.
                writer.Write("<title>");
                // OTName element can have language-dependent sub-elements; take just the body.
                var title = hit.SelectSingleNode("OTName/text()").GetSafeValue();
                writer.WriteEncodedText(title);
                writer.Write("</title>");
                // The link points to the Livelink object's viewing URL which is usually
                // the overview page with basic properties and a download link.
                writer.Write("<link>");
                var name = hit.SelectSingleNode("OTName");
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
                var summary = hit.SelectSingleNode("OTSummary");
                writer.WriteEncodedText(summary.GetSafeValue());
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
                if (type != null) {
                    writer.Write("<m:thumbnail url=\"");
                    var url = urlBase + type.GetAttribute("IconURL", "");
                    writer.WriteEncodedText(url);
                    writer.Write("\" />");
                }
                // The document size should be written in bytes to the output.
                var size = hit.SelectSingleNode("OTObjectSize");
                int length = 0;
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
