﻿#region License, Terms and Author(s)
//
// Lynclog, raw logging for Lync and Skype for business conversations
// Copyright (c) 2016 Philippe Raemy. All rights reserved.
//
//  Author(s):
//
//      Philippe Raemy
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion
using System.Xml.Linq;
using Microsoft.Lync.Model.Conversation;

namespace LyncLog
{
    class Participant : ConversationItem
    {
        string DisplayName { get; }

        public Participant(Conversation conversation, XElement xel, string displayName) : base(conversation, xel)
        {
            DisplayName = displayName;
        }
        public Participant(Conversation conversation, XElement xel) : base(conversation, xel)
        {
            DisplayName = GetAttributeValue("participant", "displayName");
        }
        public override string ToString()
        {
            return $"{base.ToString()}, {DisplayName}.";
        }

        public override string ToShortString()
        {
            return DisplayName ?? string.Empty;
        }
    }
}
