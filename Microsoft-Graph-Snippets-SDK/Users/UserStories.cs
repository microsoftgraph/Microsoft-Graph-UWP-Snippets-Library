//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Windows.Storage;

namespace Microsoft_Graph_Snippets_SDK
{
    class UserStories
    {
        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();
        private static readonly string DEFAULT_MESSAGE_BODY = "This message was sent from the Microsoft Graph SDK UWP Snippets project";
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;

        public static async Task<bool> TryUploadLargeFileAsync()
        {
            var result = await UserSnippets.UploadLargeFile();

            return result != null;
        }

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

        public static async Task<bool> TryGetMyCalendarViewAsync()
        {
            var calendar = await UserSnippets.GetMyCalendarViewAsync();
            return calendar != null;

        }

        public static async Task<bool> TryAcceptMeetingRequestAsync()
        {
            var events = await UserSnippets.GetEventsAsync();
            // Set the first event as the default meeting to accept.
            // This guarantees that we'll at least have a value to pass
            // to the snippet.
            // If the current user is the organizer of the meeting, the 
            // snippet will not work, since organizers can't accept their
            // own invitations.
            string eventToAccept = events[0].Id;

            // Look for a meeting that didn't originate with an invitation 
            // from the current user. If we can't find one, the snippet will 
            // throw an exception.
            foreach (var myEvent in events)
            {
                if (myEvent.ResponseStatus.Response != ResponseType.Organizer)
                {
                    eventToAccept = myEvent.Id;
                    break;
                }
            }
            return await UserSnippets.AcceptMeetingRequestAsync(eventToAccept);
        }

        public static async Task<bool> TryGetMessages()
        {
            var messages = await UserSnippets.GetMessagesAsync();
            return messages != null;
        }

        public static async Task<bool> TryGetMessagesThatHaveAttachments()
        {
            var messages = await UserSnippets.GetMessagesThatHaveAttachmentsAsync();
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

        public static async Task<bool> TrySendMessageWithAttachmentAsync()
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();
            var currentUser = await graphClient.Me.Request().GetAsync();

            return await UserSnippets.SendMessageWithAttachmentAsync(
                    STORY_DATA_IDENTIFIER,
                    DEFAULT_MESSAGE_BODY,
                    currentUser.UserPrincipalName
                );
        }

        // Replies to the first message returned by the GetMessages snippet
        public static async Task<bool> TryReplyToMessageAsync()
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();            
            var messages = await UserSnippets.GetMessagesAsync();
            return await UserSnippets.ReplyToMessageAsync(messages[0].Id);
        }

        // Moves a message to the "Drafts" folder
        public static async Task<bool> TryMoveMessageAsync()
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();
            var currentUser = await graphClient.Me.Request().GetAsync();

            var messages = await UserSnippets.GetMessagesAsync();
            var message = await UserSnippets.MoveMessageAsync(messages[0].Id, "Drafts");
            return message != null;
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

        // This story does not work with a personal account
        public static async Task<bool> TryGetCurrentUserPhotoStreamAsync()
        {
            Stream photoStream = await UserSnippets.GetCurrentUserPhotoStreamAsync();
            return photoStream != null;
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

        public static async Task<bool> TryGetSharingLinkAsync()
        {
            string fileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            Permission permission = await UserSnippets.GetSharingLinkAsync(fileId);
            return permission != null;
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


        //Excel stories
        public static async Task<bool> TryUploadExcelFileAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            return createdFileId != null;
        }

        public static async Task<bool> TryDeleteExcelFileAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            bool fileDeleted = await UserSnippets.DeleteExcelFileAsync(createdFileId);
            return fileDeleted;
        }

        public static async Task<bool> TryCreateExcelChartFromTableAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookChart excelWorkbookChart = await UserSnippets.CreateExcelChartFromTableAsync(createdFileId);
            return excelWorkbookChart != null;
        }

        public static async Task<bool> TryGetExcelRangeAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookRange excelWorkbookRange = await UserSnippets.GetExcelRangeAsync(createdFileId);
            return excelWorkbookRange != null;
        }

        public static async Task<bool> TryUpdateExcelRangeAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookRange excelWorkbookRange = await UserSnippets.UpdateExcelRangeAsync(createdFileId);
            return excelWorkbookRange != null;
        }

        public static async Task<bool> TryChangeExcelNumberFormatAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookRange excelWorkbookRange = await UserSnippets.ChangeExcelNumberFormatAsync(createdFileId);
            return excelWorkbookRange != null;
        }

        public static async Task<bool> TryAbsExcelFunctionAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookFunctionResult excelWorkbookFunctionResult = await UserSnippets.AbsExcelFunctionAsync(createdFileId);
            return excelWorkbookFunctionResult != null;
        }

        public static async Task<bool> TrySetExcelFormulaAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookRange excelWorkbookRange = await UserSnippets.SetExcelFormulaAsync(createdFileId);
            return excelWorkbookRange != null;
        }

        public static async Task<bool> TryAddExcelTableToUsedRangeAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookTable excelWorkbookTable = await UserSnippets.AddExcelTableToUsedRangeAsync(createdFileId);
            return excelWorkbookTable != null;
        }

        public static async Task<bool> TryAddExcelRowToTableAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            WorkbookTableRow excelWorkbookTableRow = await UserSnippets.AddExcelRowToTableAsync(createdFileId);
            return excelWorkbookTableRow != null;
        }

        public static async Task<bool> TrySortExcelTableOnFirstColumnValueAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            bool tableSorted = await UserSnippets.SortExcelTableOnFirstColumnValueAsync(createdFileId);
            return tableSorted;
        }

        public static async Task<bool> TryFilterExcelTableValuesAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            bool tableFiltered = await UserSnippets.FilterExcelTableValuesAsync(createdFileId);
            return tableFiltered;
        }

        public static async Task<bool> TryProtectExcelWorksheetAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            bool worksheetProtected = await UserSnippets.ProtectExcelWorksheetAsync(createdFileId);
            return worksheetProtected;
        }

        public static async Task<bool> TryUnprotectExcelWorksheetAsync()
        {
            string createdFileId = await UserSnippets.UploadExcelFileAsync("excelTestResource.xlsx");
            bool worksheetUnprotected = await UserSnippets.UnprotectExcelWorksheetAsync(createdFileId);
            return worksheetUnprotected;
        }
    }
}
