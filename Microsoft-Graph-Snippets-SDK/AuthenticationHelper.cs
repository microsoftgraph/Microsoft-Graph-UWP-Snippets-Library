//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using Microsoft.Graph;
using System;
using System.Diagnostics;
using System.Net.Http;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.Storage;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace Microsoft_Graph_Snippets_SDK
{
    internal static class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to Microsoft Azure Active Directory (AD).
        static string clientId = App.Current.Resources["ida:ClientID"].ToString();
        private static Uri returnUri = new Uri(App.Current.Resources["ida:ReturnUri"].ToString());
        private static readonly string CommonAuthority = App.Current.Resources["ida:AADInstance"].ToString() + @"common";
        //Property for storing the authentication context.
        public static AuthenticationContext context = new AuthenticationContext(CommonAuthority);

        public static string TokenForUser = null;
        public static DateTimeOffset expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    graphClient = new GraphServiceClient(
                        "https://microsoftgraph.chinacloudapi.cn/v1.0",
                        new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                var token = await GetTokenForUserAsync();
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                                // This header has been added to identify our sample in the Microsoft Graph service.  If extracting this code for your project please remove.
                                requestMessage.Headers.Add("SampleID", "uwp-csharp-snippets-sample");

                            }));
                    return graphClient;
                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return graphClient;
        }


        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync()
        {
            try
            {
                if (TokenForUser == null || expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {

                    AuthenticationResult result = null;

                    result = await context.AcquireTokenAsync("https://microsoftgraph.chinacloudapi.cn", clientId, returnUri);

                    TokenForUser = result.AccessToken;
                    expiration = result.ExpiresOn;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Couldn't acquire a token: " + ex.Message);
            }

            return TokenForUser;
        }


        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            graphClient = null;
            TokenForUser = null;

        }


    }

}
