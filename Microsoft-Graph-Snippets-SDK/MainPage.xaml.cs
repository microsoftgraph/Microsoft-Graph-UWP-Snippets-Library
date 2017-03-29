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
using Windows.ApplicationModel.Resources;
using Windows.ApplicationModel.Resources.Core;
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
        //Header strings for the different scope groups.
        public string ScopeGroupAll = ResourceLoader.GetForCurrentView().GetString("PersonalWorkAccess");
        public string ScopeGroupWork = ResourceLoader.GetForCurrentView().GetString("WorkAccess");
        public string ScopeGroupWorkAdmin = ResourceLoader.GetForCurrentView().GetString("WorkAdminAccess");
        public string ScopeGroupExcel = ResourceLoader.GetForCurrentView().GetString("ExcelGroup");

        public List<StoryDefinition> StoryCollection { get; private set; }
        public MainPage()
        {
            this.InitializeComponent();
            RunSelected.Label = ResourceLoader.GetForCurrentView().GetString("RunSelected");
            ClearSelection.Label = ResourceLoader.GetForCurrentView().GetString("ClearSelected");
            Disconnect.Label = ResourceLoader.GetForCurrentView().GetString("Disconnect");
            CreateStoryList();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            // Developer code - if you haven't registered the app yet, we warn you. 
            if (!App.Current.Resources.ContainsKey("ida:ClientID"))
            {
                Debug.WriteLine(ResourceLoader.GetForCurrentView().GetString("PleaseRegister"));

            }
        }
        private void CreateStoryList()
        {
            StoryCollection = new List<StoryDefinition>();

            // Stories applicable to both work or school and personal accounts
            // NOTE: All of these snippets require permissions available whether you sign into the sample 
            // with a or work or school (commercial) or personal (consumer) account

            var usersGroupName = ResourceLoader.GetForCurrentView().GetString("UsersGroup");
            var groupsGroupName = ResourceLoader.GetForCurrentView().GetString("GroupsGroup");

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("UploadLargeFile"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryUploadLargeFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetMe"),  ScopeGroup= ScopeGroupAll, RunStoryAsync = UserStories.TryGetMeAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("ReadUsers"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryGetUsersAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetDrive"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryGetCurrentUserDriveAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetEvents"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryGetEventsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("CreateEvent"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryCreateEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("UpdateEvent"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryUpdateEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("DeleteEvent"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryDeleteEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetMyCalendarView"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryGetMyCalendarViewAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("AcceptMeeting"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryAcceptMeetingRequestAsync });

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetMessages"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryGetMessages });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetMyInboxMessagesThatHaveAttachments"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryGetMessagesThatHaveAttachments });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("SendMessage"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TrySendMailAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("SendMessageWithAttachment"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TrySendMessageWithAttachmentAsync });

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetUserFiles"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryGetCurrentUserFilesAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("CreateTextFile"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryCreateFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("DownloadFile"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryDownloadFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("UpdateFile"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryUpdateFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("RenameFile"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryRenameFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("DeleteFile"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryDeleteFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("CreateFolder"), ScopeGroup = ScopeGroupAll, RunStoryAsync = UserStories.TryCreateFolderAsync });


            // Stories applicable only to work or school accounts
            // NOTE: All of these snippets will fail for lack of permissions if you sign into the sample with a personal (consumer) account

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetManager"), ScopeGroup = ScopeGroupWork, RunStoryAsync = UserStories.TryGetCurrentUserManagerAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetDirects"), ScopeGroup = ScopeGroupWork, RunStoryAsync = UserStories.TryGetDirectReportsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetPhoto"), ScopeGroup = ScopeGroupWork, RunStoryAsync = UserStories.TryGetCurrentUserPhotoAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetPhotoStream"), ScopeGroup = ScopeGroupWork, RunStoryAsync = UserStories.TryGetCurrentUserPhotoStreamAsync });

            // Stories applicable only to work or school accounts with admin access
            // NOTE: All of these snippets will fail for lack of permissions if you sign into the sample with a non-admin work account
            // or any consumer account. 

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("CreateUser"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = UserStories.TryCreateUserAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetUserGroups"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = UserStories.TryGetCurrentUserGroupsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetAllGroups"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryGetGroupsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetAGroup"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryGetGroupAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetMembers"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryGetGroupMembersAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetOwners"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryGetGroupOwnersAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("CreateGroup"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryCreateGroupAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("UpdateGroup"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryUpdateGroupAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("DeleteGroup"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryDeleteGroupAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = groupsGroupName, Title = ResourceLoader.GetForCurrentView().GetString("AddMember"), ScopeGroup = ScopeGroupWorkAdmin, RunStoryAsync = GroupStories.TryAddUserToGroup });

            // Excel snippets. These stories are applicable only to work or school accounts.
            // NOTE: All of these snippets will fail for lack of permissions if you sign into the sample with a personal (consumer) account

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("UploadXLFile"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryUploadExcelFileAsync });

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("CreateXLChart"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryCreateExcelChartFromTableAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("GetXLRange"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryGetExcelRangeAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("UpdateXLRange"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryUpdateExcelRangeAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("ChangeXLNumFormat"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryChangeExcelNumberFormatAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("XLABSFunction"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryAbsExcelFunctionAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("SetXLFormula"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TrySetExcelFormulaAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("AddXLTable"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryAddExcelTableToUsedRangeAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("AddXLTableRow"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryAddExcelRowToTableAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("SortXLTable"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TrySortExcelTableOnFirstColumnValueAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("FilterXLTable"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryFilterExcelTableValuesAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("ProtectWorksheet"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryProtectExcelWorksheetAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("UnprotectWorksheet"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryUnprotectExcelWorksheetAsync });

            StoryCollection.Add(new StoryDefinition() { GroupName = usersGroupName, Title = ResourceLoader.GetForCurrentView().GetString("DeleteXLFile"), ScopeGroup = ScopeGroupExcel, RunStoryAsync = UserStories.TryDeleteExcelFileAsync });



            var result = from story in StoryCollection group story by story.ScopeGroup into api orderby api.Key select api;
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
