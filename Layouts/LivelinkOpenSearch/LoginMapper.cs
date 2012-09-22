using System;
using System.Security.Principal;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration.Claims;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // Computes login names of Livelink users from their SharePoint counterparts that can
    // be used for user authentication or impersonation with Livelink. Users are usually
    // synchronized from SharePoint (Windows AD) to Livelink and their names are transformed
    // by using a predictable pattern.
    //
    // The following code converts MyDomain\MyUser to myuser@mycompany.com, for exaample:
    //   var mapper = new LoginMapper("{user:lc}@mycompany.com);
    //   var login = mapper.GetLoginName(SPContext.Current.Web.CurrentUser);
    //
    // The pattern describing the user name transformation is a combination of parameter
    // placeholders and other text. The parameter placeholders are enclosed in braces
    // and are replaced with values from the SharePoint user information:
    //
    //   {login}  - complete login name; with the domain if provided.
    //   {user}   - only user name; without the domain.
    //   {domain} - only the domain name; empty if no domain is provided.
    //
    // Unrecognized parameter placeholders (expressions in braces) are left intact. The parameter
    // placeholders can be used together with modifiers that transform the text value further.
    // They start with colon and are written between the parameter placeholder name and the right
    // closing brace like {domain:uc}, for example:
    //
    //   :lc - transform the value to lower-case
    //   :uc - transform the value to lower-case
    class LoginMapper {

        // Pattern to convert the user name of a SharePoint user to the user name of a Livelink
        // user. It is a combination of parameter placeholders and other text as described above.
        string Pattern { get; set; }

        // Constructor. About the parameter pattern, see the property Pattern above.
        public LoginMapper(string pattern) {
            if (pattern == null)
                throw new ArgumentNullException("pattern");
            Pattern = pattern;
        }

        // Gets the login name of the Livelink user that the specified SharePoint user maps to.
        public string GetLoginName(SPUser user) {
            if (user == null)
                throw new ArgumentNullException("user");
            // SPUser.LoginName contains domain\user for web applications with the pure Windows
            // authentication but if the claim-based authentication is used it returns an encoded
            // claim that must be decoded to the actual user login name first.
            var claim = SPClaimProviderManager.Local.ConvertSPUserToClaim(user);
            string login;
            if (SPClaimTypes.Equals(claim.ClaimType, SPClaimTypes.UserLogonName) ||
                SPClaimTypes.Equals(claim.ClaimType,
                    "http://schemas.microsoft.com/sharepoint/2009/08/claims/processidentitylogonname")) {
                login = claim.Value;
            } else if (SPClaimTypes.Equals(claim.ClaimType, SPClaimTypes.UserIdentifier) ||
                 SPClaimTypes.Equals(claim.ClaimType,
                     "http://schemas.microsoft.com/sharepoint/2009/08/claims/processidentitysid") ||
                 SPClaimTypes.Equals(claim.ClaimType,
                     "http://schemas.microsoft.com/ws/2008/06/identity/claims/primarysid")) {
                var identifier = new SecurityIdentifier(claim.Value);
                login = identifier.Translate(typeof(NTAccount)).Value;
            } else {
                throw new ApplicationException(
                    "No claim with either user name or SID was found to infer the login name from.");
            }
            // Here we assume either plain user name or a combination with the Windows domain.
            var parts = login.Split('\\');
            var name = parts.Length > 1 ? parts[1] : login;
            var domain = parts.Length > 1 ? parts[0] : "";
            return Pattern.ReplaceParameter("login", login).ReplaceParameter("user", name).
                ReplaceParameter("domain", domain);
        }
    }

    // Extends the class String with extra methods, for example:
    //
    //   // Resolves the URL template "http://{login:lc}@myhost" with login "MyUser"
    //   // to "http://myuser@myhost", for example:
    //   var url = urlTemplate.ReplaceParameter("login", login);
    //
    //   // Replaces %DOT% (case-insensitively) with the actual dot, for example:
    //   var test = "myhost%dot%com".ReplaceNC("%DOT%", () => ".") == "myhost.com";
    static class StringExtension {

        // Replaces all occurrences of the parameter placeholder in the input text
        // with the specified value. Text matching is done case-insensitively.
        // If the text to search for is not found the replace closure is never called;
        // otherwise it is called just once and its result used for all occurrences.
        public static string ReplaceParameter(this string text, string name, string value) {
            if (name == null)
                throw new ArgumentNullException("name");
            if (value == null)
                throw new ArgumentNullException("value");
            return text.ReplaceNC("{" + name + "}", () => value).
                ReplaceNC("{" + name + ":lc}", () => value.ToLower()).
                ReplaceNC("{" + name + ":uc}", () => value.ToUpper());
        }

        // Replaces all occurrences of the search parameter in the input text with the string
        // returned by the replace closure. Text matching is done case-insensitively.
        // If the text to search for is not found the replace closure is never called; otherwise
        // it is called just once and its result used for all occurrences.
        public static string ReplaceNC(this string text, string search, Func<string> replace) {
            if (search == null)
                throw new ArgumentNullException("search");
            if (replace == null)
                throw new ArgumentNullException("replace");
            // If the search string was not found just return the input text.
            var index = text.IndexOf(search, StringComparison.InvariantCultureIgnoreCase);
            if (index < 0)
                return text;
            var result = new StringBuilder();
            // The search string was found at least once; compute the replacement string.
            string value = replace();
            if (value == null)
                throw new InvalidOperationException("The replacement value was null.");
            for (var start = 0; ; ) {
                // Append the part of the input text before the current search occurrence
                // and then the replacement value.
                result.Append(text, start, index - start).Append(value);
                // Move the search start behind the just found occurrence.
                start = index + search.Length;
                // Try to find the next search occurrence.
                index = text.IndexOf(search, start, StringComparison.InvariantCultureIgnoreCase);
                // If not found append the rest of the input text from the latest search start
                // and return the result.
                if (index < 0)
                    return result.Append(text, start, text.Length - start).ToString();
            }
        }
    }
}
