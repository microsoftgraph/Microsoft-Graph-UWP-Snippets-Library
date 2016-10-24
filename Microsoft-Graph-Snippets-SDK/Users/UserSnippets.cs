//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Windows.Storage;
using Windows.Storage.Streams;
using System.Runtime.InteropServices.WindowsRuntime;

namespace Microsoft_Graph_Snippets_SDK
{
    public class UserSnippets
    {
        static string ExcelFileId = null;

        // Returns information about the signed-in user 
        public static async Task<string> GetMeAsync()
        {
            string currentUserName = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();

                var currentUserObject = await graphClient.Me.Request().GetAsync();
                currentUserName = currentUserObject.DisplayName;

                if ( currentUserName != null)
                {
                    Debug.WriteLine("Got user: " + currentUserName);
                }

            }


            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user: " + e.Error.Message);
                return null;

            }

            return currentUserName;
        }


        // Returns all of the users in the directory of the signed-in user's tenant. 
        public static async Task<IGraphServiceUsersCollectionPage> GetUsersAsync()
        {
            IGraphServiceUsersCollectionPage users = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                users = await graphClient.Users.Request().GetAsync();

                foreach ( var user in users)
                {
                    Debug.WriteLine("User: " + user.DisplayName);
                }

                return users;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get users: " + e.Error.Message);
                return null;
            }


        }

        // Creates a new user in the signed-in user's tenant. 
        // This snippet requires an admin work account.

        public static async Task<string> CreateUserAsync(string userName)
        {
            string createdUserId = null;




            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();

                //Get tenant domain from the Organization object. We use this domain to create the 
                //user's email address.
                var passwordProfile = new PasswordProfile();
                passwordProfile.Password = "pass@word1";
                var organization = await graphClient.Organization.Request().GetAsync();
                var domain = organization.CurrentPage[0].VerifiedDomains.ElementAt(0).Name;

                var user = await graphClient.Users.Request().AddAsync(new User
                {
                    AccountEnabled = true,
                    DisplayName = "User " + userName,
                    MailNickname = userName,
                    PasswordProfile = passwordProfile,
                    UserPrincipalName = userName + "@" + domain

                });

                createdUserId = user.Id;
                Debug.WriteLine("Created user: " + createdUserId);

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not create a user: " + e.Error.Message);
                return null;
            }

            return createdUserId;

        }

        // Gets the signed-in user's drive.
        public static async Task<string> GetCurrentUserDriveAsync()
        {
            string currentUserDriveId = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();

                var currentUserDrive = await graphClient.Me.Drive.Request().GetAsync();
                currentUserDriveId = currentUserDrive.Id;

                if (currentUserDriveId != null)
                {
                    Debug.WriteLine("Got user drive: " + currentUserDriveId);

                }

            }


            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user drive: " + e.Error.Message);
                return null;

            }

            return currentUserDriveId;

        }

        // Gets the signed-in user's calendar events.

        public static async Task<IUserEventsCollectionPage> GetEventsAsync()
        {
            IUserEventsCollectionPage events = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                events = await graphClient.Me.Events.Request().GetAsync();

                foreach (var myEvent in events)
                {
                    Debug.WriteLine("Got event: " + myEvent.Id);
                }

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user's events: " + e.Error.Message);
                return null;
            }

            return events;
        }

        // Creates a new event in the signed-in user's tenant.
        // Important note: This will create a user with a weak password. Consider deleting this user after you run the sample.
        public static async Task<string> CreateEventAsync()
        {
            string createdEventId = null;

            //List of attendees
            List<Attendee> attendees = new List<Attendee>();
            var attendee = new Attendee();
            var emailAddress = new EmailAddress();
            emailAddress.Address = "mara@fabrikam.com";
            attendee.EmailAddress = emailAddress;
            attendee.Type = AttendeeType.Required;
            attendees.Add(attendee);

            //Event body
            var eventBody = new ItemBody();
            eventBody.Content = "Status updates, blocking issues, and next steps";
            eventBody.ContentType = BodyType.Text;

            //Event start and end time
            var eventStartTime = new DateTimeTimeZone();
            eventStartTime.DateTime = new DateTime(2014, 12, 1, 9, 30, 0).ToString("o");
            eventStartTime.TimeZone = "UTC";
            var eventEndTime = new DateTimeTimeZone();
            eventEndTime.TimeZone = "UTC";
            eventEndTime.DateTime = new DateTime(2014, 12, 1, 10, 0, 0).ToString("o");

            //Create an event to add to the events collection

            var location = new Location();
            location.DisplayName = "Water cooler";
            var newEvent = new Event();
            newEvent.Subject = "weekly sync";
            newEvent.Location = location;
            newEvent.Attendees = attendees;
            newEvent.Body = eventBody;
            newEvent.Start = eventStartTime;
            newEvent.End = eventEndTime;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var createdEvent = await graphClient.Me.Events.Request().AddAsync(newEvent);
                createdEventId = createdEvent.Id;

                if (createdEventId != null)
                {
                    Debug.WriteLine("Created event: " + createdEventId);
                }

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not create an event: " + e.Error.Message);
                return null;
            }

            return createdEventId;

        }

        // Updates the subject of an existing event in the signed-in user's tenant.
        public static async Task<bool> UpdateEventAsync(string eventId)
        {
            bool eventUpdated = false;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var eventToUpdate = new Event();
                eventToUpdate.Subject = "Sync of the week";
                //eventToUpdate.IsAllDay = true;
                var updatedEvent = await graphClient.Me.Events[eventId].Request().UpdateAsync(eventToUpdate);

                if (updatedEvent != null)
                {
                    Debug.WriteLine("Updated event: " + eventToUpdate.Id);
                    eventUpdated = true;
                }

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not update the event: " + e.Error.Message);
                eventUpdated = false;
            }

            return eventUpdated;

        }

        // Deletes an existing event in the signed-in user's tenant.
        public static async Task<bool> DeleteEventAsync(string eventId)
        {
            bool eventDeleted = false;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var eventToDelete = await graphClient.Me.Events[eventId].Request().GetAsync();
                await graphClient.Me.Events[eventId].Request().DeleteAsync();
                Debug.WriteLine("Deleted event: " + eventToDelete.Id);
                eventDeleted = true;
            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not delete the event: " + e.Error.Message);
                eventDeleted = false;
            }

            return eventDeleted;

        }

        // Returns the first page of the signed-in user's messages.
        public static async Task<IUserMessagesCollectionPage> GetMessagesAsync()
        {
            IUserMessagesCollectionPage messages = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                messages = await graphClient.Me.Messages.Request().GetAsync();

                foreach ( var message in messages)
                {
                    Debug.WriteLine("Got message: " + message.Subject);
                }

                return messages;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get messages: " + e.Error.Message);
                return null;
            }


        }

        // Updates the subject of an existing event in the signed-in user's tenant.
        public static async Task<bool> SendMessageAsync(
            string Subject,
            string Body,
            string RecipientAddress
            )
        {
            bool emailSent = false;

            //Create recipient list
            List<Recipient> recipientList = new List<Recipient>();
            recipientList.Add(new Recipient { EmailAddress = new EmailAddress { Address = RecipientAddress.Trim() } });

            //Create message
            var email = new Message
            {
                Body = new ItemBody
                {
                    Content = Body,
                    ContentType = BodyType.Html,
                },
                Subject = Subject,
                ToRecipients = recipientList,
            };


            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                await graphClient.Me.SendMail(email, true).Request().PostAsync();
                Debug.WriteLine("Message sent");
                emailSent = true;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not send the message. The request returned this status code: " + e.Error.Message);
                emailSent = false;
            }

            return emailSent;
        }

        // Gets the signed-in user's manager. 
        // This snippet doesn't work with personal accounts.
        public static async Task<string> GetCurrentUserManagerAsync()
        {
            string currentUserManagerId = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var currentUserManager = await graphClient.Me.Manager.Request().GetAsync();
                currentUserManagerId = currentUserManager.Id;
                Debug.WriteLine("Got manager: " + currentUserManagerId);

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user's manager: " + e.Error.Message);
                return null;

            }

            return currentUserManagerId;

        }

        // Gets the signed-in user's direct reports. 
        // This snippet doesn't work with consumer accounts.

        public static async Task<IUserDirectReportsCollectionWithReferencesPage> GetDirectReportsAsync()
        {
            IUserDirectReportsCollectionWithReferencesPage directReports = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                directReports =  await graphClient.Me.DirectReports.Request().GetAsync();


                foreach (User direct in directReports)
                {
                    Debug.WriteLine("Got direct report: " + direct.DisplayName);
                }

                return directReports;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get direct reports: " + e.Error.Message);
                return null;
            }


        }


        // Gets the signed-in user's photo. 
        // This snippet doesn't work with consumer accounts.
        public static async Task<string> GetCurrentUserPhotoAsync()
        {
            string currentUserPhotoId = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var userPhoto = await graphClient.Me.Photo.Request().GetAsync();
                currentUserPhotoId = userPhoto.Id;
                Debug.WriteLine("Got user photo: " + currentUserPhotoId);

            }


            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the current user photo: " + e.Error.Message);
                return null;

            }

            return currentUserPhotoId;

        }

        // Gets the groups that the signed-in user is a member of. 
        // This snippet requires an admin work account.
        public static async Task<IUserMemberOfCollectionWithReferencesPage> GetCurrentUserGroupsAsync()
        {
            IUserMemberOfCollectionWithReferencesPage memberOfGroups = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                memberOfGroups = await graphClient.Me.MemberOf.Request().GetAsync();

                foreach (var group in memberOfGroups)
                {
                    Debug.WriteLine("Got group: " + group.Id);
                }

                return memberOfGroups;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get user groups: " + e.Error.Message);
                return null;
            }


        }


        public static async Task<IDriveItemChildrenCollectionPage> GetCurrentUserFilesAsync()
        {
            IDriveItemChildrenCollectionPage files = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                files = await graphClient.Me.Drive.Root.Children.Request().GetAsync();
                foreach (DriveItem item in files)
                {
                    Debug.WriteLine("Got file: " + item.Name);
                }

                return files;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get user files: " + e.Error.Message);
                return null;
            }


        }

        // Creates a text file in the user's root directory.
        public static async Task<string> CreateFileAsync(string fileName, string fileContent)
        {
            string createdFileId = null;
            DriveItem fileToCreate = new DriveItem();

            //Read fileContent string into a stream that gets passed as the file content
            byte[] byteArray = Encoding.ASCII.GetBytes(fileContent);
            MemoryStream fileContentStream = new MemoryStream(byteArray);

            fileToCreate.Content = fileContentStream;


            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var createdFile = await graphClient.Me.Drive.Root.ItemWithPath(fileName).Content.Request().PutAsync<DriveItem>(fileContentStream);
                createdFileId = createdFile.Id;

                Debug.WriteLine("Created file Id: " + createdFileId);


            }

            //Known bug -- file created but ServiceException "Value cannot be null" thrown
            //Workaround: catch the exception, make sure that it is the one that is expected, get the item and itemID
            catch (ServiceException se)
            {
                if (se.InnerException.Message.Contains("Value cannot be null"))
                {
                    var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                    var createdFile = await graphClient.Me.Drive.Root.ItemWithPath(fileName).Request().GetAsync();
                    createdFileId = createdFile.Id;
                    return createdFileId;

                }

                else
                {
                    Debug.WriteLine("We could not create the file. The request returned this status code: " + se.Error.Message);
                    return null;
                }

            }

            return createdFileId;
        }

        // Downloads the content of an existing file.
        public static async Task<Stream> DownloadFileAsync(string fileId)
        {
            Stream fileContent = null;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var downloadedFile = await graphClient.Me.Drive.Items[fileId].Content.Request().GetAsync();
                fileContent = downloadedFile;
                Debug.WriteLine("Downloaded file content for file: " + fileId);


            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not download the file. The request returned this status code: " + e.Error.Message);
                return null;
            }

            return fileContent;
        }

        // Adds content to a file in the user's root directory.
        public static async Task<bool> UpdateFileAsync(string fileId, string fileContent)
        {
            bool fileUpdated = false;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();

                //Read fileContent string into a stream that gets passed as the file content
                byte[] byteArray = Encoding.ASCII.GetBytes(fileContent);
                MemoryStream fileContentStream = new MemoryStream(byteArray);

                var updatedFile = await graphClient.Me.Drive.Items[fileId].Content.Request().PutAsync<DriveItem>(fileContentStream); ;

                if (updatedFile != null)
                {
                    Debug.WriteLine("Updated file Id: " + updatedFile.Id);
                    fileUpdated = true;
                }


            }
            //Known bug -- file created but ServiceException "Value cannot be null" thrown
            //Workaround: catch the exception, make sure that it is the one that is expected, return true
            catch (ServiceException se)
            {
                if (se.InnerException.Message.Contains("Value cannot be null"))
                {
                    return true;
                }

                else
                {
                    Debug.WriteLine("We could not create the file. The request returned this status code: " + se.Message);
                    return false;
                }

            }

            return fileUpdated;
        }

        // Deletes a file in the user's root directory.
        public static async Task<bool> DeleteFileAsync(string fileId)
        {
            bool fileDeleted = false;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var fileToDelete = graphClient.Me.Drive.Items[fileId].Request().GetAsync();
                await graphClient.Me.Drive.Items[fileId].Request().DeleteAsync();
                Debug.WriteLine("Deleted file Id: " + fileToDelete.Id);
                fileDeleted = true;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not delete the file. The request returned this status code: " + e.Error.Message);
                fileDeleted = false;
            }

            return fileDeleted;
        }

        // Renames a file in the user's root directory.
        public static async Task<bool> RenameFileAsync(string fileId, string newFileName)
        {
            bool fileRenamed = false;

            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();

                var fileToRename = new DriveItem();
                fileToRename.Name = newFileName;
                var renamedFile = await graphClient.Me.Drive.Items[fileId].Request().UpdateAsync(fileToRename);

                if (renamedFile != null)
                {
                    Debug.WriteLine("Renamed file Id: " + renamedFile.Id);
                    fileRenamed = true;
                }


            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not rename the file. The request returned this status code: " + e.Error.Message);
                fileRenamed = false;
            }

            return fileRenamed;
        }


        // Creates a folder in the user's root directory. 
        // This does not work with consumer accounts.
        public static async Task<string> CreateFolderAsync(string folderName)
        {
            string createFolderId = null;
            DriveItem folderToCreate = new DriveItem();
            folderToCreate.Name = folderName;
            var folder = new Folder();
            folderToCreate.Folder = folder;


            try
            {
                var graphClient = AuthenticationHelper.GetAuthenticatedClient();
                var createdFolder = await graphClient.Me.Drive.Root.Children.Request().AddAsync(folderToCreate);

                if (createdFolder != null)
                {
                    createFolderId = createdFolder.Id;
                    Debug.WriteLine("Created folder Id: " + createFolderId);

                }


            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not create the folder. The request returned this status code: " + e.Error.Message);
                return null;
            }

            return createFolderId;
        }

        //Excel snippets

        // Uploads a file and then streams the contents 
        // of an existing spreadsheet into it. You must run this at least once
        // before running the other Excel snippets. Otherwise, the other snippets
        // will fail, because they assume the presence of this file.
        // If you quit the app and run it again 5-10 minutes after uploading
        // the Excel file, the snippets will likely fail,
        // because the new file won't be indexed for search yet and the app won't
        // be able to retrieve the file Id

        public static async Task<string> UploadExcelFileAsync(string fileName)
        {
            string createdFileId = null;

            try
            {
                // First test to determine whether the file exists. Create it if it doesn't.
                // Don't bother to search if ExcelFileId variable is already populated.
                if (ExcelFileId != null)
                {
                    createdFileId = ExcelFileId;
                }
                else
                {
                    createdFileId = await SearchForFileAsync(fileName);
                }

                if (createdFileId == null)
                {

                    GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
                    DriveItem excelWorkbook = new DriveItem()
                    {
                        Name = fileName,
                        File = new Microsoft.Graph.File()
                    };

                    // Create the Excel file.

                    DriveItem excelWorkbookDriveItem = await graphClient.Me.Drive.Root.Children.Request().AddAsync(excelWorkbook);

                    createdFileId = excelWorkbookDriveItem.Id;
                    bool excelFileUploaded = await UploadExcelFileContentAsync(createdFileId);

                    if (excelFileUploaded)
                    {
                        createdFileId = excelWorkbookDriveItem.Id;
                        Debug.WriteLine("Created Excel file. Name: " + fileName + " Id: " + createdFileId);

                    }

                }
            }
            catch (ServiceException e)
            {
                Debug.WriteLine("We could not create the file: " + e.Error.Message);
            }

            // Store the file id so that you don't have to search for it
            // in subsequent snippets.
            ExcelFileId = createdFileId;
            return createdFileId;
        }

        public static async Task<string> SearchForFileAsync(string fileName)
        {
            string fileId = null;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
                // Check that this item hasn't already been created. 
                // https://graph.microsoft.io/en-us/docs/api-reference/v1.0/api/item_search
                IDriveItemSearchCollectionPage searchResults = await graphClient.Me.Drive.Root.Search(fileName).Request().GetAsync();
                foreach (var r in searchResults)
                {
                    if (r.Name == fileName)
                    {
                        fileId = r.Id;
                    }
                }
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("The search for this file failed: " + e.Error.Message);
            }

            return fileId;
        }

        
        public static async Task<bool> UploadExcelFileContentAsync(string fileId)
        {
            bool fileUploaded = false;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
                DriveItem excelDriveItem = null;

                StorageFile file = await Windows.ApplicationModel.Package.Current.InstalledLocation.GetFileAsync("excelTestResource.xlsx");
                Stream fileStream = (await file.OpenReadAsync()).AsStreamForRead();

                // Upload content to the file.
                // https://graph.microsoft.io/en-us/docs/api-reference/v1.0/api/item_uploadcontent
                excelDriveItem = await graphClient.Me.Drive.Items[fileId].Content.Request().PutAsync<DriveItem>(fileStream);

                if ( excelDriveItem != null)
                {
                    fileUploaded = true;
                }
            }
            catch (ServiceException e)
            {
                Debug.WriteLine("We failed to upload the contents of the Excel file: " + e.Error.Message);
            }

            return fileUploaded;

        }


        public static async Task<bool> DeleteExcelFileAsync(string fileId)
        {
             
            bool fileDeleted = false;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
               
                if (fileId != null)
                {

                    DriveItem w = await graphClient.Me.Drive.Items[fileId].Request().GetAsync();

                    List<Option> headers = new List<Option>()
                    {
                        new HeaderOption("if-match", "*")
                    };

                    // Delete the workbook.
                    // https://graph.microsoft.io/en-us/docs/api-reference/v1.0/api/item_delete
                    await graphClient.Me.Drive.Items[fileId].Request(headers).DeleteAsync();
                    fileDeleted = true;
                    Debug.WriteLine("Deleted the file: " + fileId);
                    ExcelFileId = null;
                }

                else
                {
                    Debug.WriteLine("We couldn't find the Excel file, so we didn't delete it.");
                }
            }


            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to delete the Excel file: " + e.Error.Message);
            }

            return fileDeleted;

        }


        public static async Task<WorkbookChart> CreateExcelChartFromTableAsync(string fileId)
        {
            WorkbookChart workbookChart = null;

            if (fileId == null)
            {
                Debug.WriteLine("Please upload the excelTestResource.xlsx file by running the Upload XL File snippet.");
                return workbookChart;
            }

            else
            {
                try
                {



                    GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                    // Get the table range.
                    WorkbookRange tableRange = await graphClient.Me.Drive.Items[fileId]
                                                               .Workbook
                                                               .Tables["CreateChartFromTable"] // Set in excelTestResource.xlsx
                                                               .Range()
                                                               .Request()
                                                               .GetAsync();

                    // Create a chart based on the table range.
                    workbookChart = await graphClient.Me.Drive.Items[fileId]
                                                                  .Workbook
                                                                  .Worksheets["CreateChartFromTable"] // Set in excelTestResource.xlsx
                                                                  .Charts
                                                                  .Add("ColumnStacked", "Auto", tableRange.Address)
                                                                  .Request()
                                                                  .PostAsync();

                    Debug.WriteLine("Created the Excel chart: " + workbookChart.Name);
                }
                catch (Microsoft.Graph.ServiceException e)
                {
                    Debug.WriteLine("We failed to create the Excel chart: " + e.Error.Message);
                }

                return workbookChart;
            }
        }


        public static async Task<WorkbookRange> GetExcelRangeAsync(string fileId)
        {
            WorkbookRange workbookRange = null;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
                // GET https://graph.microsoft.com/beta/me/drive/items/012KW42LDENXUUPCMYQJDYX3CLZMORQKGT/workbook/worksheets/Sheet1/Range(address='A1')
                workbookRange = await graphClient.Me.Drive.Items[fileId]
                                                              .Workbook
                                                              .Worksheets["GetUpdateRange"] // Set in excelTestResource.xlsx
                                                              .Range("A1")
                                                              .Request()
                                                              .GetAsync();

                Debug.WriteLine("Got the Excel workbook range: " + workbookRange.Address);

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to get the Excel workbook range: " + e.Error.Message);
            }

            return workbookRange;
        }


        public static async Task<WorkbookRange> UpdateExcelRangeAsync(string fileId)
        {
            WorkbookRange workbookRange = null;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
                // GET https://graph.microsoft.com/beta/me/drive/items/012KW42LDENXUUPCMYQJDYX3CLZMORQKGT/workbook/worksheets/Sheet1/Range(address='A1')
                WorkbookRange rangeToUpdate = await graphClient.Me.Drive.Items[fileId]
                                                              .Workbook
                                                              .Worksheets["GetUpdateRange"] // Set in excelTestResource.xlsx
                                                              .Range("A1")
                                                              .Request()
                                                              .GetAsync();

                // Forming the JSON for the updated values
                JArray arr = rangeToUpdate.Values as JArray;
                JArray arrInner = arr[0] as JArray;
                arrInner[0] = $"{arrInner[0] + "Updated"}"; // JToken

                // Create a dummy WorkbookRange object so that we only PATCH the values we want to update.
                WorkbookRange dummyWorkbookRange = new WorkbookRange();
                dummyWorkbookRange.Values = arr;

                // Update the range values.
                workbookRange = await graphClient.Me.Drive.Items[fileId]
                                                              .Workbook
                                                              .Worksheets["GetUpdateRange"] // Set in excelTestResource.xlsx
                                                              .Range("A1")
                                                              .Request()
                                                              .PatchAsync(dummyWorkbookRange);

                Debug.WriteLine("Updated the Excel workbook range: " + workbookRange.Address);

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to get the Excel workbook range: " + e.Error.Message);
            }

            return workbookRange;
        }

        public static async Task<WorkbookRange> ChangeExcelNumberFormatAsync(string fileId)
        {
            WorkbookRange workbookRange = null;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();
                string excelWorksheetId = "ChangeNumberFormat";
                string rangeAddress = "E2";

                // Forming the JSON for 
                JArray arr = JArray.Parse(@"[['$#,##0.00;[Red]$#,##0.00']]"); // Currency format

                WorkbookRange dummyWorkbookRange = new WorkbookRange();
                dummyWorkbookRange.NumberFormat = arr;


                workbookRange = await graphClient.Me.Drive.Items[fileId]
                                                              .Workbook
                                                              .Worksheets[excelWorksheetId]
                                                              .Range(rangeAddress)
                                                              .Request()
                                                              .PatchAsync(dummyWorkbookRange);

                Debug.WriteLine("Updated the number format of the Excel workbook range: " + workbookRange.Address);
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to change the number format: " + e.Error.Message);
            }
            return workbookRange;
        }

        public static async Task<WorkbookFunctionResult> AbsExcelFunctionAsync(string fileId)
        {
            WorkbookFunctionResult workbookFunctionResult = null;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Get the absolute value of -10
                JToken inputNumber = JToken.Parse("-10");

                workbookFunctionResult = await graphClient.Me.Drive.Items[fileId].Workbook.Functions.Abs(inputNumber).Request().PostAsync();
                Debug.WriteLine("Ran the Excel ABS function: " + workbookFunctionResult.Value);

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to run the ABS function: " + e.Error.Message);
            }

            return workbookFunctionResult;
        }

        public static async Task<WorkbookRange> SetExcelFormulaAsync(string fileId)
        {
            WorkbookRange workbookRange = null;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Forming the JSON for updating the formula
                var arr = JArray.Parse(@"[['=A4*B4']]");

                // We want to use a dummy workbook object so that we only send the property we want to update.
                var dummyWorkbookRange = new WorkbookRange();
                dummyWorkbookRange.Formulas = arr;

                workbookRange = await graphClient.Me.Drive.Items[fileId]
                                                              .Workbook
                                                              .Worksheets["SetFormula"]
                                                              .Range("C4")
                                                              .Request()
                                                              .PatchAsync(dummyWorkbookRange);

                Debug.WriteLine("Set an Excel formula in this workbook range: " + workbookRange.Address);
            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to set the formula: " + e.Error.Message);
            }
            return workbookRange;
        }

        public static async Task<WorkbookTable> AddExcelTableToUsedRangeAsync(string fileId)
        {
            WorkbookTable workbookTable = null;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Get the used range of this worksheet. This results in a call to the service.
                WorkbookRange workbookRange = await graphClient.Me.Drive.Items[fileId]
                                                              .Workbook
                                                              .Worksheets["AddTableUsedRange"]
                                                              .UsedRange()
                                                              .Request()
                                                              .GetAsync();


                // Create the dummy workbook object. Must use the AdditionalData property for this. 
                WorkbookTable dummyWorkbookTable = new WorkbookTable();
                Dictionary<string, object> requiredPropsCreatingTableFromRange = new Dictionary<string, object>();
                requiredPropsCreatingTableFromRange.Add("address", workbookRange.Address);
                requiredPropsCreatingTableFromRange.Add("hasHeaders", false);
                dummyWorkbookTable.AdditionalData = requiredPropsCreatingTableFromRange;

                // Create a table based on the address of the workbookRange. 
                // This results in a call to the service.
                // https://graph.microsoft.io/en-us/docs/api-reference/v1.0/api/tablecollection_add
                workbookTable = await graphClient.Me.Drive.Items[fileId]
                                                              .Workbook
                                                              .Worksheets["AddTableUsedRange"]
                                                              .Tables
                                                              .Add(false, workbookRange.Address)
                                                              .Request()
                                                              .PostAsync();

                Debug.WriteLine("Added this Excel table to the used range: " + workbookTable.Name);

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to add the table to the used range: " + e.Error.Message);
            }

            return workbookTable;
        }

        public static async Task<WorkbookTableRow> AddExcelRowToTableAsync(string fileId)
        {
            WorkbookTableRow workbookTableRow = null;

            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Create the table row to insert. This assumes that the table has 2 columns.
                // You'll want to make sure you give a JSON array that matches the size of the table.
                WorkbookTableRow newWorkbookTableRow = new WorkbookTableRow();
                newWorkbookTableRow.Index = 0;
                JArray myArr = JArray.Parse("[[\"ValueA2\",\"ValueA3\"]]");
                newWorkbookTableRow.Values = myArr;

                //// Insert a new row. This results in a call to the service.
                workbookTableRow = await graphClient.Me.Drive.Items[fileId]
                                                                 .Workbook
                                                                 .Tables["Table1"]
                                                                 .Rows
                                                                 .Request()
                                                                 .AddAsync(newWorkbookTableRow);

                Debug.WriteLine("Added a row to Table 1.");

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to add the table row: " + e.Error.Message);
            }

            return workbookTableRow;
        }

        public static async Task<bool> SortExcelTableOnFirstColumnValueAsync(string fileId)
        {
            bool tableSorted = false;

            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Create the sorting options.
                WorkbookSortField sortField = new WorkbookSortField()
                {
                    Ascending = true,
                    SortOn = "Value",
                    Key = 0
                };

                List<WorkbookSortField> workbookSortFields = new List<WorkbookSortField>() { sortField };

                // Sort the table. This results in a call to the service.
                await graphClient.Me.Drive.Items[fileId].Workbook.Tables["Table2"]
                                                                          .Sort
                                                                          .Apply(true, "", workbookSortFields)
                                                                          .Request()
                                                                          .PostAsync();

                tableSorted = true;
                Debug.WriteLine("Sorted the table on this value: " + sortField.SortOn);

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to sort the table: " + e.Error.Message);
            }

            return tableSorted;
        }

        public static async Task<bool> FilterExcelTableValuesAsync(string fileId)
        {
            bool tableFiltered = false;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Filter the table. This results in a call to the service.
                await graphClient.Me.Drive.Items[fileId]
                                          .Workbook
                                          .Tables["FilterTableValues"]
                                          .Columns["1"] // This is a one based index.
                                          .Filter
                                          .ApplyValuesFilter(JArray.Parse("[\"2\"]"))
                                          .Request()
                                          .PostAsync();
                tableFiltered = true;
                Debug.WriteLine("Filtered the table.");

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to filter the table: " + e.Error.Message);
            }

            return tableFiltered;
        }

        public static async Task<bool> ProtectExcelWorksheetAsync(string fileId)
        {
            bool worksheetProtected = false;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Protect the worksheet.
                await graphClient.Me.Drive.Items[fileId]
                                          .Workbook
                                          .Worksheets["ProtectWorksheet"]
                                          .Protection
                                          .Protect()
                                          .Request()
                                          .PostAsync();

                worksheetProtected = true;
                Debug.WriteLine("Protected the worksheet.");

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to protect the worksheet: " + e.Error.Message);
            }

            return worksheetProtected;
        }

        public static async Task<bool> UnprotectExcelWorksheetAsync(string fileId)
        {
            bool worksheetUnProtected = false;
            try
            {
                GraphServiceClient graphClient = AuthenticationHelper.GetAuthenticatedClient();

                // Protect the worksheet.
                await graphClient.Me.Drive.Items[fileId]
                                          .Workbook
                                          .Worksheets["ProtectWorksheet"]
                                          .Protection
                                          .Unprotect()
                                          .Request()
                                          .PostAsync();

                worksheetUnProtected = true;
                Debug.WriteLine("Unprotected the worksheet.");

            }
            catch (Microsoft.Graph.ServiceException e)
            {
                Debug.WriteLine("We failed to unprotect the worksheet: " + e.Error.Message);
            }

            return worksheetUnProtected;
        }

    }
}
