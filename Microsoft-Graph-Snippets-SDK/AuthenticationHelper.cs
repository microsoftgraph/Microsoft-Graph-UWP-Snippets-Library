//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

//using Microsoft.Graph.Authentication;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
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

namespace Microsoft_Graph_Snippets_SDK
{
    internal static class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to Microsoft Azure Active Directory (AD).
        static string clientId = App.Current.Resources["ida:ClientID"].ToString();
        static string returnUrl = App.Current.Resources["ida:ReturnUrl"].ToString();
        static string authString = "https://login.microsoftonline.com/common";
        static string resourceUrl = "https://graph.microsoft.com/";

        public static string TokenForUser;
        public static DateTimeOffset expiration;

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static GraphServiceClient GetAuthenticatedClientAsync()
        {
            if (graphClient == null)
            {
                //*********************************************************************
                // setup Microsoft Graph Client for user...
                //*********************************************************************
                try
                {
                    graphClient = new GraphServiceClient(
                        "https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                var token = await GetTokenForUserAsync();
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

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
            if (TokenForUser == null || expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
            {
                var redirectUri = new Uri(returnUrl);
                string authority = authString;
                AuthenticationContext authenticationContext = new AuthenticationContext(authority, false);
                AuthenticationResult userAuthnResult = await authenticationContext.AcquireTokenAsync(resourceUrl,
                    clientId, redirectUri, PromptBehavior.RefreshSession);
                TokenForUser = userAuthnResult.AccessToken;
                expiration = userAuthnResult.ExpiresOn;
            }

            return TokenForUser;
        }


        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            graphClient = null;

        }


    }
}
