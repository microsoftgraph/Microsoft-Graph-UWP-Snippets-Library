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

namespace Microsoft_Graph_Snippets_SDK
{
    class UserSnippets
    {

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
            string createdUserName = null;




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

                createdUserName = user.DisplayName;
                Debug.WriteLine("Created user: " + createdUserName);

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not create a user: " + e.Error.Message);
                return null;
            }

            return createdUserName;

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

    }
}
