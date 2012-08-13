using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Web;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // Makes HTTP requests against the Livelink server. The typical usage:
    //
    //   var client = new WebClient("htttp://myserver/livelink/llisapi.dll");
    //   // SSO has not ben turned on for the client - authenticate explicitly
    //   // with the system administrator credentials, for example.
    //   using (var credentials = new Credentials()) {
    //       "Admin".ForEach(ch => credentials.Name.AppendChar(ch));
    //       "livelink".ForEach(ch => credentials.Password.AppendChar(ch));
    //       client.Authenticate(credentials);
    //   }
    //   var request = client.CreateRequest(query);
    //   using (var content = request.GetResponseContent()) {
    //       ...
    //   }
    class WebClient {

        // URL base made of Protocol, host and path to access Livelink:
        // http://myserver/livelink/llisapi.dll, for example.
        public Uri Url { get; private set; }

        // If set to true the current SharePoint credentials will be passed
        // to the HTTP request with no cookie authentication.
        public bool UseSSO { get; set; }

        // Cookie container shared for all HTTP requests to support authentication.
        CookieContainer Cookies { get; set; }

        // Constructor. About the parameter url, see the property Url above.
        public WebClient(string url) {
            if (url == null)
                throw new ArgumentNullException("url");
            Url = new Uri(url);
            Cookies = new CookieContainer();
        }

        // Creates a new HTTP request with the specified URL query part. Just set
        // the HTTP method and go. Have look at the extension GetResponseContent too.
        // If UseSSO has not been set to true you must call the Authenticate method
        // before creating any request by this method.
        public HttpWebRequest CreateRequest(string query) {
            if (query == null)
                throw new ArgumentNullException("query");
            var request = WebRequest.CreateHttp(Url + "?" + query);
            request.UserAgent = UserAgent;
            request.CookieContainer = Cookies;
            if (UseSSO)
                request.Credentials = CredentialCache.DefaultNetworkCredentials;
            else
                if (Cookies.Count == 0)
                    throw new InvalidOperationException(
                        "The HTTP client has not been authenticated yet.");
            return request;
        }

        // Authenticates future HTTP requests explicitly by issung one (func=ll.login)
        // that will retrieve the authentication cookie for the user with the specified
        // credentials and store it for the future requests. This methods needs not be
        // called if you set UseSSO to true.
        public void Authenticate(Credentials credentials) {
            if (credentials == null)
                throw new ArgumentNullException("credentials");
            var request = WebRequest.CreateHttp(Url);
            request.UserAgent = UserAgent;
            request.CookieContainer = Cookies;
            // Function ll.login accepts only POST requests of the login form. It is better
            // to conceal the credentials which are sent over the network. In addition, you
            // should use the HTTPS protocol to be secure against sniffing attacks.
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            using (var input = request.GetRequestStream())
            using (var writer = new StreamWriter(input, new UTF8Encoding(false))) {
                writer.Write("func=ll.login&CurrentClientTime=");
                var time = DateTime.Now.ToString("D/yyyy/MM/dd:HH:mm:ss",
                    CultureInfo.InvariantCulture);
                writer.Write(HttpUtility.UrlEncode(time));
                // The NextURL parameter is mandatory.
                writer.Write("&NextURL=");
                writer.Write(HttpUtility.UrlEncode(Url.AbsolutePath));
                writer.Write("&UserName=");
                writer.Write(HttpUtility.UrlEncode(credentials.Name.ToInsecureString()));
                writer.Write("&Password=");
                writer.Write(HttpUtility.UrlEncode(credentials.Password.ToInsecureString()));
            }
            using (var response = (HttpWebResponse)request.GetResponse()) {
                // The actual output of the login request is ignored. Just the cookies are wanted.
                if (response.StatusCode != HttpStatusCode.OK)
                    throw new ApplicationException(string.Format(
                        "HTTP request to {0} failed: {1} ({2}).",
                        response.StatusDescription, response.StatusCode));
                if (Cookies.Count == 0)
                    throw new ApplicationException("Authentication cookie was not received.");
            }
        }

        // Common user agent: the product name from the assembly attributes.
        public static string UserAgent {
            get {
                return Assembly.GetExecutingAssembly().
                    GetCustomAttribute<AssemblyProductAttribute>().Product;
            }
        }
    }

    // Extends the class HttpWebRequest with extra methods, for example:
    //
    //   var request = WebRequest.CreateHttp("htttp://myserver/livelink/llisapi.dll");
    //   using (var content = request.GetResponseContent()) {
    //       ...
    //   }
    static class HttpWebRequestExtension {

        // Convenience method for HTTP GET requests returning the response stream of the just
        // created request object. The stream must be disposed when not needed anymore.
        public static Stream GetResponseContent(this HttpWebRequest request) {
            request.Method = "GET";
            request.AllowReadStreamBuffering = false;
            var response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
                return response.GetResponseStream();
            response.Dispose();
            throw new ApplicationException(string.Format("HTTP request to {0} failed: {1} ({2}).",
                response.ResponseUri, response.StatusDescription, response.StatusCode));
        }
    }
}
