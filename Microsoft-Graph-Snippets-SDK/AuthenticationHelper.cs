//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using Microsoft.Graph.Authentication;
using Microsoft.Graph;
using System;
using System.Diagnostics;
using System.Net.Http;
using System.Linq;
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

        private static GraphServiceClient graphClient = null;

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static async Task<GraphServiceClient> GetAuthenticatedClientAsync()
        {
            if (graphClient == null)
            {
                var authenticationProvider = new OAuth2AuthenticationProvider(
                        clientId,
                        returnUrl,
                        new string[]
                        {
                        "openid",
                        "offline_access",
                        "https://graph.microsoft.com/User.Read",
                        //"https://graph.microsoft.com/User.ReadWrite",
                        "https://graph.microsoft.com/Mail.Send",
                        //"https://graph.microsoft.com/User.ReadBasic.All",
                        "https://graph.microsoft.com/Calendars.ReadWrite",
                        "https://graph.microsoft.com/Mail.ReadWrite",
                        //"https://graph.microsoft.com/User.ReadWrite.All",
                        //"https://graph.microsoft.com/Group.ReadWrite.All",
                        "https://graph.microsoft.com/Files.ReadWrite",
                        //"https://graph.microsoft.com/Directory.AccessAsUser.All"
                    });

                await authenticationProvider.AuthenticateAsync();

                graphClient = new GraphServiceClient(authenticationProvider);
            }

            return graphClient;
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
