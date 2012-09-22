using System;
using System.Globalization;
using System.Web;
using System.Web.UI;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // Generates an OpenSearch descriptor (OSDX) returns it as the repsonse content.
    // If the operation fails the page returns the HTTP status code 500 with the error
    // message in the status description.
    //
    // For example, the following parameters will make the URL below:
    //   LivelinkUrl       - http://myserver/livelink/llisapi.dll
    //   TargetAppID       - myserver
    //   LoginPattern      - {user:lc}
    //   ExtraParams       - lookfor1=allwords&fullTextMode=allwords&hhterms=true
    //   MaxSummaryLength  - 185
    //   IgnoreSSLWarnings - true
    //
    //   http://sharepoint/_layouts/LivelinkOpenSearch/GetOSDX.aspx?livelinkUrl=http%3A%2F%2F
    //   myserver%2Flivelink%2Fllisapi.dll&targetAppID=myserver&loginPattern=%7Buser:lc%7D&
    //   extraParams=lookfor1%3Dallwords%26fullTextMode%3Dallwords&%26hhterms%3Dtrue&
    //   maxSummaryLength=185&ignoreSSLWarnings=true
    //
    // Notice that some parameters are URL parts and should be URL encoded so that this page
    // gets them correctly when it is called.
    public partial class GetOSDXPage : SearchConnectorPageBase {

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

        // Errors occurring during the search are reported by HTTP error code 500 by default.
        // This parameter makes the error message be returned as single search hit for OpenSearch
        // clients that do not show HTTP errors to the user.
        bool ReportErrorAsHit { get; set; }

        // Reads and checks the provided URL parameters, formats the OpenSearch descriptor
        // and writes it to the response output as text/xml. If anything fails the HTTP status
        // code 500 will be returned in the response with the error message in the status
        // description.
        protected override void Render(HtmlTextWriter writer) {
            try {
                ParseUrl();
                Response.ContentType = "text/xml";
                FormatContent(writer);
            } catch (Exception exception) {
                Response.StatusCode = 500;
                Response.StatusDescription = exception.Message;
            }
        }

        // Sets values of properties in this class from the provided URL parameters
        // or from the assumed parameter defaults.
        void ParseUrl() {
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
            ExtraParams = Request.QueryString["extraParams"];
            var maxSummaryLength = Request.QueryString["maxSummaryLength"];
            if (!string.IsNullOrEmpty(maxSummaryLength))
                MaxSummaryLength = int.Parse(maxSummaryLength, CultureInfo.InvariantCulture);
            ReportErrorAsHit = "true".Equals(Request.QueryString["reportErrorAsHit"],
                StringComparison.InvariantCultureIgnoreCase);
        }

        // Writes the OSDX content to the response output in the XML format.
        void FormatContent(HtmlTextWriter writer) {
            var livelinkUri = new Uri(LivelinkUrl);
            var authentication = UseSSO ? "useSSO=true" : string.Format(
                "targetAppID={0}&loginPattern={1}", HttpUtility.UrlEncode(TargetAppID),
                HttpUtility.UrlEncode(LoginPattern));
            var limit = MaxSummaryLength > 0 ? string.Format("maxSummaryLength={0}&",
                MaxSummaryLength) : "";
            var error = ReportErrorAsHit > 0 ? "reportErrorAsHit=true&" : "";
            var certification = IgnoreSSLWarnings ? "ignoreSSLWarnings=true&" : "";
            var urlTemplate = string.Format(
                "{0}/ExecuteQuery.aspx?query={{searchTerms}}&livelinkUrl={1}&" +
                "{2}&count={{count}}&startIndex={{startIndex}}&extraParams=" +
                "{3}&{4}{5}{6}inputEncoding={{inputEncoding}}&" +
                "outputEncoding={{outputEncoding}}&language={{language}}",
                PageUrlPath, HttpUtility.UrlEncode(LivelinkUrl),
                authentication, HttpUtility.UrlEncode(ExtraParams), limit, error, certification);
            var urlOSDX = string.Format("{0}/GetOSDX.aspx?livelinkUrl={1}&" +
                "extraParams={2}&{3}{4}{5}{6}", PageUrlPath, HttpUtility.UrlEncode(LivelinkUrl),
                HttpUtility.UrlEncode(ExtraParams), limit, error, certification, authentication);
            writer.Write(string.Format(@"<OpenSearchDescription xmlns=""http://a9.com/-/spec/opensearch/1.1/"">
  <ShortName>Search Enterprise at {0}</ShortName>
  <LongName>Search Livelink Enterprise Workspace at {0}</LongName>
  <InternalName xmlns=""http://schemas.microsoft.com/Search/2007/location"">search_{0}</InternalName>
  <Description>Searches content in the Enterprise Workspace of the Livelink server at {0}.</Description>
  <Image height=""32"" width=""32"" type=""image/png"">{1}/img/style/images/app_content_server32_b8.png</Image>
  <Url type=""application/rss+xml"" rel=""results"" template=""{2}""/>
  <Url type=""text/html"" template=""{2}&amp;format=html""/>
  <Url type=""application/opensearchdescription+xml"" rel=""self"" template=""{3}""/>
  <Tags>Livelink OpenSearch</Tags>
  <Query role=""example"" searchTerms=""livelink"" />
  <AdultContent>false</AdultContent>
  <Developer>Ferdinand Prantl</Developer>
  <Contact>prantlf@gmail.com</Contact>
  <Attribution>Copyright (c) 2012 Ferdinand Prantl, All rights reserved.</Attribution>
  <SyndicationRight>Open</SyndicationRight>
  <InputEncoding>UTF-8</InputEncoding>
  <OutputEncoding>UTF-8</OutputEncoding>
  <Language>*</Language>
  
  <ms-ose:ResultsProcessing format=""application/rss+xml"" xmlns:ms-ose=""http://schemas.microsoft.com/opensearchext/2009/"">
    <ms-ose:PropertyDefaultValues>
      <ms-ose:Property schema=""http://schemas.microsoft.com/windows/2008/propertynamespace"" name=""System.PropList.ContentViewModeForSearch"">prop:~System.ItemNameDisplay;System.LayoutPattern.PlaceHolder;~System.ItemPathDisplay;~System.Search.AutoSummary;System.LayoutPattern.PlaceHolder;System.LayoutPattern.PlaceHolder;System.LayoutPattern.PlaceHolder</ms-ose:Property>
    </ms-ose:PropertyDefaultValues>
  </ms-ose:ResultsProcessing>
</OpenSearchDescription>", HttpUtility.HtmlEncode(livelinkUri.Host),
                HttpUtility.HtmlEncode(livelinkUri.GetComponents(UriComponents.SchemeAndServer,
                    UriFormat.Unescaped)), HttpUtility.HtmlEncode(urlTemplate),
                    HttpUtility.HtmlEncode(urlOSDX)));
        }
    }
}
