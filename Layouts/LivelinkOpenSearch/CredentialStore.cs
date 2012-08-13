using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using Microsoft.BusinessData.Infrastructure.SecureStore;
using Microsoft.Office.SecureStoreService.Server;
using Microsoft.SharePoint;

namespace LivelinkSearchConnector.Layouts.LivelinkOpenSearch {

    // Wraps user name and password to a single secured object. Typical usages:
    //
    //   using (var credentials = Credentials()) {
    //       "Admin".ForEach(ch => credentials.Name.AppendChar(ch));
    //       "livelink".ForEach(ch => credentials.Password.AppendChar(ch));
    //       ...
    //   }
    //
    //   // Name and password will be made read-only by this constructor. The new
    //   // object will own the secured strings; make copies if you dispose them
    //   // and the credentials object is still supposed to exist.
    //   using (object owning the name and password values) {
    //       ...
    //       return new Credentials(name.Copy(), password.Copy()));
    //   }
    sealed class Credentials : IDisposable {

        // Name of the user with these credentials. Even the name should be secret.
        public SecureString Name { get; private set; }

        // Password of the user with these credentials.
        public SecureString Password { get; private set; }

        // Constructor. Name and password values will be empty and will be modifiable
        // in the created object. You can make them read-only later. They will be owned
        // and disposed with it.
        public Credentials() {
            Name = new SecureString();
            Password = new SecureString();
        }

        // Constructor. Name and password values are made read-only when stored
        // in the created object. They will be owned and disposed with it.
        public Credentials(SecureString name, SecureString password) {
            if (name == null)
                throw new ArgumentNullException("name");
            if (password == null)
                throw new ArgumentNullException("password");
            Name = name;
            Password = password;
            // If someone prepared the name and password before they should not be
            // tampered with after the creation of this credentials wrapper.
            Name.MakeReadOnly();
            Password.MakeReadOnly();
        }

        // Securely clears the memory used by the credentials. It should be always
        // called as soon as they are not necessary.
        public void Dispose() {
            Name.Dispose();
            Password.Dispose();
        }
    }

    // Utilizes the SharePoint Secure Store Service to access credentials. Typical usage
    // to get credentials of a master account accessible not for the current SharePoint
    // user but for the application pool account (ReverToSelf authentication type):
    //
    // Some SharePoint site is needed for the context. In the following example we use
    // the current site but we want to impersonate the service running the code.
    // var store = new CredentialStore(SPContext.Current.Site, true);
    // using (var credentials = store.GetCredentials(TargetAppID)) {
    //     ...
    // }
    class CredentialStore {

        // Contextual SharePoint site to access the Secure Store Service with.
        SPSite ContextSite { get; set; }

        // If the service executing the code should be used to authenticate the Secure Store
        // access instead of the current SharePoint user. It is possible only with the full trust.
        bool RevertToSelf { get; set; }

        // Constructor. About the parameters, see the respective properties above.
        public CredentialStore(SPSite contextSite, bool revertToSelf) {
            if (contextSite == null)
                throw new ArgumentNullException("contextSite");
            ContextSite = contextSite;
            RevertToSelf = revertToSelf;
        }

        // Gets credentials for the specified target application. It decides about
        // the authenticated context according to the constructor parameters.
        public Credentials GetCredentials(string targetAppID) {
            if (targetAppID == null)
                throw new ArgumentNullException("targetAppID");
            if (!RevertToSelf)
                return GetCredentials(ContextSite, targetAppID);
            Credentials credentials = null;
            SPSecurity.RunWithElevatedPrivileges(delegate {
                using (var contextSite = new SPSite(ContextSite.ID))
                    credentials = GetCredentials(contextSite, targetAppID);
            });
            return credentials;
        }

        // Delegated getter of credentials for the specified target application that
        // uses the contextSite as the carrier of the authenticated context. The target
        // application must contain non-empty UserName and Password credentials.
        Credentials GetCredentials(SPSite contextSite, string targetAppID) {
            var provider = SecureStoreProviderFactory.Create();
            var context = (ISecureStoreServiceContext)provider;
            context.Context = SPServiceContext.GetContext(contextSite);
            using (var credentials = provider.GetCredentials(targetAppID)) {
                if (credentials == null)
                    throw new ApplicationException(
                        "Credentials for the target application were null.");
                var name = credentials.FirstOrDefault(item =>
                    item.CredentialType == SecureStoreCredentialType.UserName);
                if (name == null || name.Credential.Length == 0)
                    throw new ApplicationException(
                        "User name credential for the target application was empty.");
                var password = credentials.FirstOrDefault(item =>
                    item.CredentialType == SecureStoreCredentialType.Password);
                if (password == null || password.Credential.Length == 0)
                    throw new ApplicationException(
                        "Password credential for the target application was empty.");
                return new Credentials(name.Credential.Copy(), password.Credential.Copy());
            }
        }
    }

    // Extends the class SecureString with extra methods, for example:
    //
    //   using (var password = new SecureString()) {
    //       "livelink".ForEach(ch => password.AppendChar(ch));
    //       Console.WriteLine(password.ToInsecureString())
    //   }
    static class SecureStringExtension {

        // Returns the otherwise inaccessible value of the SecureString as the plain
        // System.String that can be used by methods not accepting the secureed object.
        public static string ToInsecureString(this SecureString input) {
            IntPtr pointer = Marshal.SecureStringToBSTR(input);
            try {
                return Marshal.PtrToStringBSTR(pointer);
            } finally {
                Marshal.FreeBSTR(pointer);
            }
        }
    }
}
