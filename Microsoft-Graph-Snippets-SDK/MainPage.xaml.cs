//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.


using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Microsoft_Graph_Snippets_SDK
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public List<StoryDefinition> StoryCollection { get; private set; }
        public MainPage()
        {
            this.InitializeComponent();
            CreateStoryList();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            // Developer code - if you haven't registered the app yet, we warn you. 
            if (!App.Current.Resources.ContainsKey("ida:ClientID"))
            {
                Debug.WriteLine("Oops - App not registered with Office 365. To run this sample, you must register it with Office 365. See Readme for more info.");

            }
        }
        private void CreateStoryList()
        {
            StoryCollection = new List<StoryDefinition>();

            // These stories require your app to have permission to access your organization's directory. 
            // Comment them if you're not going to run the app with that permission level.

            // User stories

            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get Me", RunStoryAsync = UserStories.TryGetMeAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Read Users", RunStoryAsync = UserStories.TryGetUsersAsync });

            // Comment the TryCreateUserAsync story if you're not running this sample with an admin work account.
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Create User", RunStoryAsync = UserStories.TryCreateUserAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get Drive", RunStoryAsync = UserStories.TryGetCurrentUserDriveAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get Events", RunStoryAsync = UserStories.TryGetEventsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Create Event", RunStoryAsync = UserStories.TryCreateEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Update Event", RunStoryAsync = UserStories.TryUpdateEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Delete Event", RunStoryAsync = UserStories.TryDeleteEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get Messages", RunStoryAsync = UserStories.TryGetMessages });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Send Message", RunStoryAsync = UserStories.TrySendMailAsync });

            // Comment the TryGetCurrentUserManagerAsync, TryGetDirectReportsAsync, and TryGetCurrentUserPhotoAsync stories if you're running this sample with a consumer account.
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get Manager", RunStoryAsync = UserStories.TryGetCurrentUserManagerAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get Directs", RunStoryAsync = UserStories.TryGetDirectReportsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get Photo", RunStoryAsync = UserStories.TryGetCurrentUserPhotoAsync });

            // Comment the TryGetCurrentUserGroupsAsync story if you're not running this sample with an admin work account.
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get User Groups", RunStoryAsync = UserStories.TryGetCurrentUserGroupsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Get User Files", RunStoryAsync = UserStories.TryGetCurrentUserFilesAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Create Text File", RunStoryAsync = UserStories.TryCreateFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Download File", RunStoryAsync = UserStories.TryDownloadFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Update File", RunStoryAsync = UserStories.TryUpdateFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Rename File", RunStoryAsync = UserStories.TryRenameFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Delete File", RunStoryAsync = UserStories.TryDeleteFileAsync });

            // Comment the TryCreateFolderAsync story if you're running this sample with a consumer account.
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users", Title = "Create Folder", RunStoryAsync = UserStories.TryCreateFolderAsync });


            // Group stories
            // NOTE: All of these snippets will fail for lack of permissions if you're running the sample with a non-admin work account
            // or any consumer account. Comment these lines if you're not using an admin work account.
            StoryCollection.Add(new StoryDefinition() { GroupName = "Groups", Title = "Get All Groups", RunStoryAsync = GroupStories.TryGetGroupsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Groups", Title = "Get a Group", RunStoryAsync = GroupStories.TryGetGroupAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Groups", Title = "Get Members", RunStoryAsync = GroupStories.TryGetGroupMembersAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Groups", Title = "Get Owners", RunStoryAsync = GroupStories.TryGetGroupOwnersAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Groups", Title = "Create Group", RunStoryAsync = GroupStories.TryCreateGroupAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Groups", Title = "Update Group", RunStoryAsync = GroupStories.TryUpdateGroupAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Groups", Title = "Delete Group", RunStoryAsync = GroupStories.TryDeleteGroupAsync });


            // Organization stories

            //StoryCollection.Add(new StoryDefinition() { GroupName = "Organization", Title = "Get Org Drives", RunStoryAsync = OrganizationStories.TryGetDrivesAsync });


            var result = from story in StoryCollection group story by story.GroupName into api orderby api.Key select api;
            StoriesByApi.Source = result;
        }


        private async void RunSelectedStories_Click(object sender, RoutedEventArgs e)
        {

            await runSelectedAsync();
        }

        private async Task runSelectedAsync()
        {
            ResetStories();
            Stopwatch sw = new Stopwatch();

            foreach (var story in StoryGrid.SelectedItems)
            {
                StoryDefinition currentStory = story as StoryDefinition;
                currentStory.IsRunning = true;
                sw.Restart();
                bool result = false;
                try
                {
                    result = await currentStory.RunStoryAsync();
                    Debug.WriteLine(String.Format("{0}.{1} {2}", currentStory.GroupName, currentStory.Title, (result) ? "passed" : "failed"));
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("{0}.{1} failed. Exception: {2}", currentStory.GroupName, currentStory.Title, ex.Message);
                    result = false;

                }
                currentStory.Result = result;
                sw.Stop();
                currentStory.DurationMS = sw.ElapsedMilliseconds;
                currentStory.IsRunning = false;


            }

            // To shut down this app when the Stories complete, uncomment the following line. 
            // Application.Current.Exit();
        }

        private void ResetStories()
        {
            foreach (var story in StoryCollection)
            {
                story.Result = null;
                story.DurationMS = null;
            }
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            StoryGrid.SelectedItems.Clear();
        }

        private void Disconnect_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationHelper.SignOut();
            StoryGrid.SelectedItems.Clear();
        }

    }
}
