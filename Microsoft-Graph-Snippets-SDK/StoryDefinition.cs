//Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
//See LICENSE in the project root for license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microsoft_Graph_Snippets_SDK
{

    //All of the properties of a story.
    public class StoryDefinition : ViewModelBase
    {
        //The name of the group to which the story belongs.
        public string GroupName { get; set; }
        //Display title of the story.
        public string Title { get; set; }

        //The permission scope group to which the story belongs.
        public string ScopeGroup { get; set; }

        // Delegate method to call.
        public Func<Task<bool>> RunStoryAsync { get; set; }

        //Specifies whether the story is running.
        bool _isRunning = false;
        public bool IsRunning
        {
            get
            {
                return _isRunning;
            }
            set
            {
                SetProperty(ref _isRunning, value);
            }
        }

        //Success or failure result.
        bool? _result = null;
        public bool? Result
        {
            get
            {
                return _result;
            }
            set
            {
                SetProperty(ref _result, value);
            }
        }

        //Length of time the story takes to run.
        long? _durationMS = 0;
        public long? DurationMS
        {
            get
            {
                return _durationMS;
            }
            set
            {
                SetProperty(ref _durationMS, value);
            }
        }

    }


}

