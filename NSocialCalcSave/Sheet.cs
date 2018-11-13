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
    using System.Collections.Generic;

    public interface ISheet
    {
        ICollection<KeyValuePair<string, ICell>> Cells { get; }
        ICollection<KeyValuePair<string, string>> ColWidths { get; }
        ICollection<KeyValuePair<string, bool>> ColHides { get; }
        ICollection<KeyValuePair<int, int>> RowHeights { get; }
        ICollection<KeyValuePair<int, bool>> RowHides { get; }
        ICollection<INamedRange> Names { get; }
        ICollection<KeyValuePair<int, string>> Layouts { get; }
        ICollection<KeyValuePair<int, string>> Fonts { get; }
        ICollection<KeyValuePair<int, string>> Colors { get; }
        ICollection<KeyValuePair<int, string>> BorderStyles { get; }
        ICollection<KeyValuePair<int, string>> CellFormats { get; }
        ICollection<KeyValuePair<int, string>> ValueFormats { get; }
        int LastCol { get; }
        int LastRow { get; }
        string DefaultColWidth { get; }
        int DefaultRowHeight { get; }
        int DefaultTextFormat { get; }
        int DefaultNonTextFormat { get; }
        int DefaultLayout { get; }
        int DefaultFont { get; }
        int DefaultTextValueFormat { get; }
        int DefaultNonTextValueFormat { get; }
        int DefaultColor { get; }
        int DefaultBgColor { get; }
        string CircularReferenceCell { get; }
        string Recalc { get; }
        bool NeedsRecalc { get; }
        int UserMaxCol { get; }
        int UserMaxRow { get; }
        string CopiedFrom { get; }
    }

    public sealed class Sheet : ISheet
    {
        public ICollection<KeyValuePair<string, ICell>> Cells { get; set; }

        public ICollection<KeyValuePair<string, string>> ColWidths { get; set; }
        public ICollection<KeyValuePair<string, bool>> ColHides { get; set; }
        public ICollection<KeyValuePair<int, int>> RowHeights { get; set; }
        public ICollection<KeyValuePair<int, bool>> RowHides { get; set; }
        public ICollection<INamedRange> Names { get; set; }
        public ICollection<KeyValuePair<int, string>> Layouts { get; set; }
        public ICollection<KeyValuePair<int, string>> Fonts { get; set; }
        public ICollection<KeyValuePair<int, string>> Colors { get; set; }
        public ICollection<KeyValuePair<int, string>> BorderStyles { get; set; }
        public ICollection<KeyValuePair<int, string>> CellFormats { get; set; }
        public ICollection<KeyValuePair<int, string>> ValueFormats { get; set; }

        public int LastCol { get; set; }
        public int LastRow { get; set; }
        public string DefaultColWidth { get; set; }
        public int DefaultRowHeight { get; set; }
        public int DefaultTextFormat { get; set; }
        public int DefaultNonTextFormat { get; set; }
        public int DefaultLayout { get; set; }
        public int DefaultFont { get; set; }
        public int DefaultTextValueFormat { get; set; }
        public int DefaultNonTextValueFormat { get; set; }
        public int DefaultColor { get; set; }
        public int DefaultBgColor { get; set; }
        public string CircularReferenceCell { get; set; }
        public string Recalc { get; set; }
        public bool NeedsRecalc { get; set; }
        public int UserMaxCol { get; set; }
        public int UserMaxRow { get; set; }

        public string CopiedFrom { get; set; }
    }
}
