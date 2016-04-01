//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage;


namespace Microsoft_Graph_Snippets_SDK
{
    class GroupStories
    {

        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();

        public static async Task<bool> TryGetGroupsAsync()
        {
            var groups = await GroupSnippets.GetGroupsAsync();
            return groups != null;
        }

        public static async Task<bool> TryGetGroupAsync()
        {
            var groups = await GroupSnippets.GetGroupsAsync();
            var groupId = groups[0].Id;
            var groupName = await GroupSnippets.GetGroupAsync(groupId);
            return groupName != null;
        }

        public static async Task<bool> TryGetGroupMembersAsync()
        {
            // Get all groups, then pass the first group id from the list.
            var groups = await GroupSnippets.GetGroupsAsync();
            var groupId = groups[0].Id;
            var members = await GroupSnippets.GetGroupMembersAsync(groupId);
            return members != null;
        }

        public static async Task<bool> TryGetGroupOwnersAsync()
        {
            // Get all groups, then pass the first group id from the list.
            var groups = await GroupSnippets.GetGroupsAsync();
            var groupId = groups[0].Id;
            var members = await GroupSnippets.GetGroupOwnersAsync(groupId);
            return members != null;
        }

        public static async Task<bool> TryCreateGroupAsync()
        {
            string createdGroup = await GroupSnippets.CreateGroupAsync(STORY_DATA_IDENTIFIER);
            return createdGroup != null;
        }

        public static async Task<bool> TryUpdateGroupAsync()
        {
            // Create an group first, then update it.
            string createdGroup = await GroupSnippets.CreateGroupAsync(STORY_DATA_IDENTIFIER);
            return await GroupSnippets.UpdateGroupAsync(createdGroup);
        }

        public static async Task<bool> TryDeleteGroupAsync()
        {
            // Create a group first, then delete it.
            string createdGroup = await GroupSnippets.CreateGroupAsync(STORY_DATA_IDENTIFIER);
            return await GroupSnippets.DeleteGroupAsync(createdGroup);
        }

    }
}

