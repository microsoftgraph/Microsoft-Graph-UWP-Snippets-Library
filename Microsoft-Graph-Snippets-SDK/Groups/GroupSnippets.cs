//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Microsoft_Graph_Snippets_SDK
{
    class GroupSnippets
    {

        // Returns all of the groups in your tenant's directory.
        public static async Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync()
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();
            
            try
            {

                var groups = await graphClient.Groups.Request().GetAsync();

                foreach (var group in groups )
                {

                    Debug.WriteLine("Got group: " + group.DisplayName);
                }

                return groups;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get groups: " + e.Message);
                return null;
            }

        }


        // Returns the display name of a specific group.
        public static async Task<string> GetGroupAsync(string groupId)
        {
            string groupName = null;
            //JObject jResult = null;
            var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();
            try
            {
                var group = await graphClient.Groups[groupId].Request().GetAsync();
                groupName = group.DisplayName;
                Debug.WriteLine("Got group: " + groupName);

            }


            catch (Exception e)
            {
                Debug.WriteLine("We could not get the specified group: " + e.Message);
                return null;

            }

            return groupName;
        }

        public static async Task<IGroupMembersCollectionWithReferencesPage> GetGroupMembersAsync(string groupId)
        {
            IGroupMembersCollectionWithReferencesPage members = null;

            var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();

            try
            {
                var group = await graphClient.Groups[groupId].Request().GetAsync();
                members = group.Members;

                //Work around the fact that some groups don't have owners.
                //Return a non-null collection whenever the call has worked
                //but the group has no owners.
                if (members != null)
                {
                    foreach (var member in members)
                    {
                        Debug.WriteLine("Member Id:" + member.Id);

                    }
                }

                else
                {
                    members = new GroupMembersCollectionWithReferencesPage();
                }
            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get the group members: " + e.Message);
                return null;
            }

            return members;

        }

        public static async Task<IGroupOwnersCollectionWithReferencesPage> GetGroupOwnersAsync(string groupId)
        {
            IGroupOwnersCollectionWithReferencesPage owners = null;
            var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();

            try
            {

                var group = await graphClient.Groups[groupId].Request().GetAsync();
                owners = group.Owners;


                //Work around the fact that some groups don't have owners.
                //Return a non-null collection whenever the call has worked
                //but the group has no owners.
                if (owners != null)
                {
                    foreach (var owner in owners)
                    {
                        Debug.WriteLine("Owner Id:" + owner.Id);
                    }
                }
                else
                {
                    owners = new GroupOwnersCollectionWithReferencesPage();
                } 

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get the group owners: " + e.Message);
                return null;
            }

            return owners;

        }


        // Creates a new security group in the tenant.
        public static async Task<string> CreateGroupAsync(string groupName)
        {
            //JObject jResult = null;
            string createdGroupId = null;
            var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();


            try
            {
                var group = await graphClient.Groups.Request().AddAsync(new Group
                {
                    GroupTypes = new List<string> { "Unified" },
                    DisplayName = groupName,
                    Description = "This group was created by the snippets app.",
                    MailNickname = groupName,
                    MailEnabled = false,
                    SecurityEnabled = true
                });

                createdGroupId = group.Id;

                Debug.WriteLine("Created group:" + createdGroupId);

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not create a group: " + e.Message);
                return null;
            }

            return createdGroupId;

        }


        // Updates the description of an existing group.
        public static async Task<bool> UpdateGroupAsync(string groupId)
        {
            bool groupUpdated = false;
            var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();

            try
            {
                var groupToUpdate = new Group();
                groupToUpdate.Description = "This group was updated by the snippets app.";
                var updatedGroup = await graphClient.Groups[groupId].Request().UpdateAsync(groupToUpdate);

                //Possible bug; REST call returns 204 (no content), so UpdateAsync doesn't return a group. 
                //The UpdateAsync() method signature indicates that a group should be returned.
                //if (updatedGroup != null)
                //{
                    //Debug.WriteLine("Updated group:" + updatedGroup.Id);
                    groupUpdated = true;
                //}

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not update the group: " + e.Message);
                groupUpdated = false;
            }

            return groupUpdated;

        }


        // Deletes an existing group in the tenant.
        public static async Task<bool> DeleteGroupAsync(string groupId)
        {
            bool eventDeleted = false;
            var graphClient = AuthenticationHelper.GetAuthenticatedClientAsync();

            try
            {
                var groupToDelete = await graphClient.Groups[groupId].Request().GetAsync();
                await graphClient.Groups[groupId].Request().DeleteAsync();
                Debug.WriteLine("Deleted group:" + groupToDelete.DisplayName);
                eventDeleted = true;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not delete the group: " + e.Message);
                eventDeleted = false;
            }

            return eventDeleted;

        }

    }
}

