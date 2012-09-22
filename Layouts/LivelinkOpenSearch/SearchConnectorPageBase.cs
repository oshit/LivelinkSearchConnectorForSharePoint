using System;
using Microsoft.SharePoint.WebControls;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // Base class for pages in this solution.
    public class SearchConnectorPageBase : LayoutsPageBase  {

        // Returns the starting part of this page URL: scheme, host, port and path without file
        // name; the last part of the path will be the parent folder of the file.
        protected string PageUrlPath {
            get {
                var url = Request.Url.GetComponents(UriComponents.SchemeAndServer |
                    UriComponents.Path, UriFormat.Unescaped);
                var slash = url.LastIndexOf('/');
                return url.Substring(0, slash);
            }
        }
    }
}
