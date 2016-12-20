#region (C) 2016 Atif Aziz. Portions (C) 2008, 2009, 2010 Socialtext, Inc.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

namespace NSocialCalcSave
{
    public interface INamedRange
    {
        string Name { get; }
        string Description { get; }
        string Definition { get; }
    }

    public sealed class NamedRange : INamedRange
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Definition { get; set; }
    }
}