//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage;

namespace Microsoft_Graph_Snippets_SDK
{
    class UserStories
    {
        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();
        private static readonly string DEFAULT_MESSAGE_BODY = "This message was sent from the Microsoft Graph SDK UWP Snippets project";
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }

        public static async Task<bool> TryGetUsersAsync()
        {
            var users = await UserSnippets.GetUsersAsync();
            return users != null;
        }

        // This story requires an admin work account.

        public static async Task<bool> TryCreateUserAsync()
        {
            string createdUser = await UserSnippets.CreateUserAsync(STORY_DATA_IDENTIFIER);
            return createdUser != null;
        }

        public static async Task<bool> TryGetCurrentUserDriveAsync()
        {
            string driveId = await UserSnippets.GetCurrentUserDriveAsync();
            return driveId != null;
        }

        public static async Task<bool> TryGetEventsAsync()
        {
            var events = await UserSnippets.GetEventsAsync();
            return events != null;
        }

        public static async Task<bool> TryCreateEventAsync()
        {
            string createdEvent = await UserSnippets.CreateEventAsync();
            return createdEvent != null;
        }

        public static async Task<bool> TryUpdateEventAsync()
        {
            // Create an event first, then update it.
            string createdEvent = await UserSnippets.CreateEventAsync();
            return await UserSnippets.UpdateEventAsync(createdEvent);
        }

        public static async Task<bool> TryDeleteEventAsync()
        {
            // Create an event first, then delete it.
            string createdEvent = await UserSnippets.CreateEventAsync();
            return await UserSnippets.DeleteEventAsync(createdEvent);
        }

        public static async Task<bool> TryGetMessages()
        {
            var messages = await UserSnippets.GetMessagesAsync();
            return messages != null;
        }

        public static async Task<bool> TrySendMailAsync()
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();
            var currentUser = await graphClient.Me.Request().GetAsync();
            return await UserSnippets.SendMessageAsync(
                    STORY_DATA_IDENTIFIER,
                    DEFAULT_MESSAGE_BODY,
                    currentUser.UserPrincipalName
                );
        }

        // This story does not work with a personal account
        public static async Task<bool> TryGetCurrentUserManagerAsync()
        {
            string managerName = await UserSnippets.GetCurrentUserManagerAsync();
            return managerName != null;
        }

        // This story does not work with a personal account
        public static async Task<bool> TryGetDirectReportsAsync()
        {
            var users = await UserSnippets.GetDirectReportsAsync();
            return users != null;
        }

        // This story does not work with a personal account
        public static async Task<bool> TryGetCurrentUserPhotoAsync()
        {
            string photoId = await UserSnippets.GetCurrentUserPhotoAsync();
            return photoId != null;
        }

        // This story requires an admin work account.
        public static async Task<bool> TryGetCurrentUserGroupsAsync()
        {
            var groups = await UserSnippets.GetCurrentUserGroupsAsync();
            return groups != null;
        }

        public static async Task<bool> TryGetCurrentUserFilesAsync()
        {
            var files = await UserSnippets.GetCurrentUserFilesAsync();
            return files != null;
        }

        public static async Task<bool> TryCreateFileAsync()
        {
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return createdFileId != null;
        }

        public static async Task<bool> TryDownloadFileAsync()
        {
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            var fileContent = await UserSnippets.DownloadFileAsync(createdFileId);
            return fileContent != null;
        }

        public static async Task<bool> TryUpdateFileAsync()
        {
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return await UserSnippets.UpdateFileAsync(createdFileId, STORY_DATA_IDENTIFIER);
        }

        public static async Task<bool> TryRenameFileAsync()
        {
            var newFileName = Guid.NewGuid().ToString();
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return await UserSnippets.RenameFileAsync(createdFileId, newFileName);
        }

        public static async Task<bool> TryDeleteFileAsync()
        {
            var fileName = Guid.NewGuid().ToString();
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return await UserSnippets.DeleteFileAsync(createdFileId);
        }

        // This story does not work with a consumer account.
        public static async Task<bool> TryCreateFolderAsync()
        {
            var createdFolderId = await UserSnippets.CreateFolderAsync(Guid.NewGuid().ToString());
            return createdFolderId != null;
        }
    }
}
