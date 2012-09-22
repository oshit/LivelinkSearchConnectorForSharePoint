using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Web;
using System.Web.UI;
using Microsoft.SharePoint;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // The page can return either XML compliant with OpenSearch 1.1 to be integrated
    // with OpenSearch clients or HTML displayable in the browser, which can be used
    // for testing purposes or for OpenSearch clients that display HTML.
    enum OutputFormat { XML, HTML }

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
    //   StartIndex        - 0
    //   Count             - 10
    //   ExtraParams       - lookfor1=allwords&fullTextMode=allwords&hhterms=true
    //   MaxSummaryLength  - 185
    //   IgnoreSSLWarnings - true
    //
    //   http://sharepoint/_layouts/LivelinkOpenSearch/ExecuteQuery.aspx?query=XML+Search&
    //   livelinkUrl=http%3A%2F%2Fmyserver%2Flivelink%2Fllisapi.dll&targetAppID=myserver&
    //   loginPattern=%7Buser:lc%7D&startIndex=0&count=10&extraParams=lookfor1%3Dallwords
    //   %26fullTextMode%3Dallwords&%26hhterms%3Dtrue&maxSummaryLength=185&
    //   ignoreSSLWarnings=true
    //
    // Notice that some parameters are URL parts and should be URL encoded so that this page
    // gets them correctly when it is called. The parameter LoginPattern is URL encoded too;
    // some expressions in braces could be mistakenly resolved by the search federator before
    // it calls the URL.
    public partial class ExecuteQueryPage : SearchConnectorPageBase {

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

        // Start index of the first search result to return. If not provided the number 0
        // (the first page) assumed; to start with the very first search hit available.
        int StartIndex { get; set; }

        // The maximum count of returned search results. If not provided the actual count
        // will be decided by the search engine. Usually not all available hits are returned;
        // the maximum count can be limited to 1000 from peformance reasons, for example.
        int Count { get; set; }

        // Additional URL query string to append to the Livelink XML Search API URL.
        // The string should not start with ampersand and will be appended as-is.
        // For example, the value "lookfor1=allwords&hhterms=true" requires returning only those
        // results where all entered search terms can be found and enables highlighting the search
        // terms in the document summary displayed together with the search hit.
        string ExtraParams { get; set; }

        // Can limit the maximum length of the textual summary that is displayed below a search
        // hit to give a hint what is the document about. It usually contains highlighted
        // search terms. I saw Livelink returning really long passages (more than 1300 characters)
        // while I have not seen SharePoint returning more than 185 characters. It can make the
        // result list long and difficult to browse. You can trim the length of the displayed
        // summary by this parameter to never exceed the specified length.
        int MaxSummaryLength { get; set; }

        // The encoding of the input search terms. It is used to check that only either ASCII
        // or UTF-8 URL parameters were sent. Other encodings are not supported now.
        string InputEncoding { get; set; }

        // The expected encoding of the search results. It is used to check that only either
        // ASCII or UTF-8 output was requested. Other encodings are not supported now.
        string OutputEncoding { get; set; }

        // The preferred language for the search results. It is currently unused.
        string Language { get; set; }

        // The search results output format; XML or HTML. This controls also the MIME type
        // of the reponse: text/xml or text/html.
        OutputFormat Format { get; set; }

        // Errors occurring during the search are reported by HTTP error code 500 by default.
        // This parameter makes the error message be returned as single search hit for OpenSearch
        // clients that do not show HTTP errors to the user.
        bool ReportErrorAsHit { get; set; }

        // Reads and checks the provided URL parameters, performs the Livelink XML Search
        // query, transforms its results to RSS 2.0 and writes the content to the response
        // output as text/xml. If anything fails the HTTP status code 500 will be returned
        // in the response with the error message in the status description.
        protected override void Render(HtmlTextWriter writer) {
            string searchUrl = null;
            try {
                ParseUrl();
                if (IgnoreSSLWarnings)
                    ServicePointManager.ServerCertificateValidationCallback =
                        (sender, cert, chain, sslError) => true;
                var query = GetQuery();
                searchUrl = LivelinkUrl + "?" + SearchHelper.ConvertToBrowserUsage(query);
                using (var results = GetResults(query)) {
                    Transformer transformer;
                    switch (Format) {
                        case OutputFormat.XML:
                            transformer = new XMLTransformer(Query, searchUrl, DescriptorUrl)
                                { MaxSummaryLength = MaxSummaryLength };
                            Response.ContentType = "text/xml";
                            break;
                        case OutputFormat.HTML:
                            var previousUrl = StartIndex > 0 && Count > 0 ?
                                GetOtherPageUrl(false) : null;
                            var nextUrl = Count > 0 ? GetOtherPageUrl(true) : null;
                            transformer = new HTMLTransformer(Query, searchUrl,
                                    previousUrl, nextUrl, DescriptorUrl)
                                { MaxSummaryLength = MaxSummaryLength };
                            Response.ContentType = "text/html";
                            break;
                        default:
                            throw new InvalidOperationException(string.Format(
                                "Unexpected output format: {0}.", Format));
                    }
                    transformer.TransformResults(results, writer);
                }
            } catch (Exception exception) {
                // If parsing the reponse format failed the machine processing will be
                // supported better by returning a HTTP error. The browser response will
                // not be so friendly but because the output format is parsed at the very
                // beginning this should not be a problem. HTTP success code should not be
                // returned unless we are sure that the scenario is interactive and the user
                // will be able to see the reponse in the browser.
                if (Format == OutputFormat.HTML) {
                    Response.ContentType = "text/html";
                    FormatHtmlError(exception, writer);
                } else if (ReportErrorAsHit) {
                    FormatXmlError(exception, searchUrl, writer);
                } else {
                    Response.StatusCode = 500;
                    Response.StatusDescription = exception.Message;
                }
            }
        }

        // Sets values of properties in this class from the provided URL parameters
        // or from the assumed parameter defaults.
        void ParseUrl() {
            // Repsonse formt is deduced at the very start to be able to format errors
            // that may occur in this method already.
            var format = Request.QueryString["format"];
            if (!string.IsNullOrEmpty(format))
                Format = (OutputFormat) Enum.Parse(typeof(OutputFormat), format, true);
            ReportErrorAsHit = "true".Equals(Request.QueryString["reportErrorAsHit"],
                StringComparison.InvariantCultureIgnoreCase);
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
                StartIndex = int.Parse(startIndex, CultureInfo.InvariantCulture);
            var count = Request.QueryString["count"];
            if (!string.IsNullOrEmpty(count))
                Count = int.Parse(count, CultureInfo.InvariantCulture);
            ExtraParams = Request.QueryString["extraParams"];
            var maxSummaryLength = Request.QueryString["maxSummaryLength"];
            if (!string.IsNullOrEmpty(maxSummaryLength))
                MaxSummaryLength = int.Parse(maxSummaryLength, CultureInfo.InvariantCulture);
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

        // Returns URL that would navigate the previous or the next page of search results
        // relatively to the currently rendered one.
        string GetOtherPageUrl(bool next) {
            var startIndex = StartIndex;
            if (next)
                startIndex += Count;
            else
                startIndex -= Count;
            var url = Request.Url.ToString();
            var start = url.IndexOf("&startIndex=", StringComparison.InvariantCultureIgnoreCase);
            if (start < 0)
                start = url.IndexOf("?startIndex=", StringComparison.InvariantCultureIgnoreCase);
            if (start < 0)
                return null;
            start += 12;
            var end = url.IndexOf("&", start, StringComparison.InvariantCultureIgnoreCase);
            url = end > 0 ? url.Remove(start, end - start) : url = url.Remove(start);
            return url.Insert(start, startIndex.ToString(CultureInfo.InvariantCulture));
        }

        // Writes information about the error to the response output in the XML/RSS format.
        void FormatXmlError(Exception exception, string searchUrl, HtmlTextWriter writer) {
            var host = searchUrl != null ? HttpUtility.HtmlEncode(new Uri(searchUrl).Host) :
                Request.Url.Host;
            var link = searchUrl != null ? HttpUtility.HtmlEncode(searchUrl) :
                HttpUtility.HtmlEncode(Request.Url.AbsoluteUri);
            writer.Write(string.Format(@"<?xml version=""1.0"" encoding=""UTF-8""?>
 <rss version=""2.0"" xmlns:openSearch=""http://a9.com/-/spec/opensearch/1.1/"">
   <channel>
     <title>Livelink Enterprise at {0}</title>
     <link>{1}</link>
     <description>Search results of the query executed against the Enterprise Workspace of the Livelink server at {0}.</description>
     <openSearch:totalResults>1</openSearch:totalResults>
     <openSearch:startIndex>1</openSearch:startIndex>
     <openSearch:itemsPerPage>1</openSearch:itemsPerPage>
     <item>
       <title>{2}</title>
       <description>{3}</description>
     </item>
   </channel>
 </rss>", host, link, HttpUtility.HtmlEncode(exception.Message),
                HttpUtility.HtmlEncode(exception.ToString())));
        }


        // Writes information about the error to the response output in the HTML format.
        void FormatHtmlError(Exception exception, HtmlTextWriter writer) {
            writer.WriteFullBeginTag("html");
            // Make the title of the page in the browser caption a short constant.
            writer.WriteFullBeginTag("head");
            writer.WriteFullBeginTag("title");
            writer.WriteEncodedText("Error");
            writer.WriteEndTag("title");
            writer.WriteEndTag("head");
            // Print a title, the topmost exception message and the complete stacktrace below.
            writer.WriteFullBeginTag("body");
            writer.WriteFullBeginTag("h3");
            writer.WriteEncodedText("The Search Failed");
            writer.WriteEndTag("h3");
            writer.WriteFullBeginTag("p");
            writer.WriteEncodedText(exception.Message);
            writer.WriteEndTag("p");
            writer.WriteFullBeginTag("p");
            writer.WriteFullBeginTag("small");
            var stackTrace = exception.ToString().Replace(Environment.NewLine, "<br>");
            writer.WriteEncodedText(stackTrace);
            writer.WriteEndTag("small");
            writer.WriteEndTag("p");
            writer.WriteEndTag("body");
            writer.WriteEndTag("html");
        }

        // Returns the URL of the OpenSearch descriptor (OSDX) file.
        string DescriptorUrl {
            get {
                var livelinkUri = new Uri(LivelinkUrl);
                var authentication = UseSSO ? "useSSO=true" : string.Format(
                    "targetAppID={0}&loginPattern={1}", HttpUtility.UrlEncode(TargetAppID),
                    HttpUtility.UrlEncode(LoginPattern));
                var limit = MaxSummaryLength > 0 ? string.Format("maxSummaryLength={0}&",
                    MaxSummaryLength) : "";
                var error = ReportErrorAsHit > 0 ? "reportErrorAsHit=true&" : "";
                var certification = IgnoreSSLWarnings ? "ignoreSSLWarnings=true&" : "";
                return string.Format(
                    "{0}/GetOSDX.aspx?livelinkUrl={1}&extraParams={2}&{3}{4}{5}{6}",
                    PageUrlPath, HttpUtility.UrlEncode(LivelinkUrl),
                    HttpUtility.UrlEncode(ExtraParams), limit, error, certification,
                    authentication);
            }
        }
    }
}
