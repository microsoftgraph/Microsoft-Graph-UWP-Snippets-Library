﻿//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.


using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Graph;

// NOTE: All groups snippets work only with admin work accounts.


namespace Microsoft_Graph_Snippets_SDK
{
    class GroupSnippets
    {

        // Returns all of the groups in your tenant's directory.
        public static async Task<IGraphServiceGroupsCollectionPage> GetGroupsAsync()
        {
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();
            
            try
            {

                var groups = await graphClient.Groups.Request().GetAsync();

                foreach (var group in groups )
                {

                    Debug.WriteLine("Got group: " + group.DisplayName);
                }

                return groups;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get groups: " + e.Error.Message);
                return null;
            }

        }


        // Returns the display name of a specific group.
        public static async Task<string> GetGroupAsync(string groupId)
        {
            string groupName = null;
            //JObject jResult = null;
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();
            try
            {
                var group = await graphClient.Groups[groupId].Request().GetAsync();
                groupName = group.DisplayName;
                Debug.WriteLine("Got group: " + groupName);

            }


            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the specified group: " + e.Error.Message);
                return null;

            }

            return groupName;
        }

        public static async Task<IGroupMembersCollectionWithReferencesPage> GetGroupMembersAsync(string groupId)
        {
            IGroupMembersCollectionWithReferencesPage members = null;

            var graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {
                var group = await graphClient.Groups[groupId].Request().Expand("members").GetAsync();
                members = group.Members;


                foreach (var member in members)
                {
                    Debug.WriteLine("Member Id:" + member.Id);

                }

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the group members: " + e.Error.Message);
                return null;
            }

            return members;

        }

        public static async Task<IGroupOwnersCollectionWithReferencesPage> GetGroupOwnersAsync(string groupId)
        {
            IGroupOwnersCollectionWithReferencesPage owners = null;
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {

                var group = await graphClient.Groups[groupId].Request().Expand("owners").GetAsync();
                owners = group.Owners;

                foreach (var owner in owners)
                {
                    Debug.WriteLine("Owner Id:" + owner.Id);
                }

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not get the group owners: " + e.Error.Message);
                return null;
            }

            return owners;

        }


        // Creates a new group in the tenant.
        public static async Task<string> CreateGroupAsync(string groupName)
        {
            //JObject jResult = null;
            string createdGroupId = null;
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();


            try
            {
                var group = await graphClient.Groups.Request().AddAsync(new Group
                {
                    GroupTypes = new List<string> { "Unified" },
                    DisplayName = groupName,
                    Description = "This group was created by the snippets app.",
                    MailNickname = groupName,
                    MailEnabled = false,
                    SecurityEnabled = false
                });

                createdGroupId = group.Id;

                Debug.WriteLine("Created group:" + createdGroupId);

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not create a group: " + e.Error.Message);
                return null;
            }

            return createdGroupId;

        }


        // Updates the description of an existing group.
        public static async Task<bool> UpdateGroupAsync(string groupId)
        {
            bool groupUpdated = false;
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {
                var groupToUpdate = new Group();
                groupToUpdate.Description = "This group was updated by the snippets app.";

                //The underlying REST call returns a 204 (no content), so we can't retrieve the updated group
                //with this call.
                await graphClient.Groups[groupId].Request().UpdateAsync(groupToUpdate);
                Debug.WriteLine("Updated group:" + groupId);
                groupUpdated = true;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not update the group: " + e.Error.Message);
                groupUpdated = false;
            }

            return groupUpdated;

        }


        // Deletes an existing group in the tenant.
        public static async Task<bool> DeleteGroupAsync(string groupId)
        {
            bool eventDeleted = false;
            var graphClient = AuthenticationHelper.GetAuthenticatedClient();

            try
            {
                var groupToDelete = await graphClient.Groups[groupId].Request().GetAsync();
                await graphClient.Groups[groupId].Request().DeleteAsync();
                Debug.WriteLine("Deleted group:" + groupToDelete.Id);
                eventDeleted = true;

            }

            catch (ServiceException e)
            {
                Debug.WriteLine("We could not delete the group: " + e.Error.Message);
                eventDeleted = false;
            }

            return eventDeleted;

        }

    }
}

