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
using Microsoft.Identity.Client;

namespace Microsoft_Graph_Snippets_SDK
{
    internal static class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to Microsoft Azure Active Directory (AD).
        static string clientId = App.Current.Resources["ida:ClientID"].ToString();
        static string returnUrl = App.Current.Resources["ida:ReturnUrl"].ToString();


        public static PublicClientApplication PublicClientApp = null;
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
                var scopes = new string[]
                        {
                        "https://graph.microsoft.com/User.Read",
                        "https://graph.microsoft.com/User.ReadWrite",
                        "https://graph.microsoft.com/User.ReadBasic.All",
                        "https://graph.microsoft.com/Mail.Send",
                        "https://graph.microsoft.com/Calendars.ReadWrite",
                        "https://graph.microsoft.com/Mail.ReadWrite",
                        "https://graph.microsoft.com/Files.ReadWrite",

                        // Admin-only scopes. Comment these out if you're running the sample with a non-admin work account.
                        // You won't be able to sign in with a non-admin work account if you request these scopes.
                        // These scopes will be ignored if you leave them uncommented and run the sample with a consumer account.
                        // See the MainPage.xaml.cs file for all of the operations that won't work if you're not running the 
                        // sample with an admin work account.
                        "https://graph.microsoft.com/Directory.AccessAsUser.All",
                        "https://graph.microsoft.com/User.ReadWrite.All",
                        "https://graph.microsoft.com/Group.ReadWrite.All",


                    };

                PublicClientApp = new PublicClientApplication(clientId);
                AuthenticationResult authResult = await PublicClientApp.AcquireTokenAsync(scopes);

                TokenForUser = authResult.Token;
                expiration = authResult.ExpiresOn;
            }

            return TokenForUser;
        }


        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            foreach (var user in PublicClientApp.Users)
            {
                user.SignOut();
            }
            graphClient = null;

        }


    }
}
