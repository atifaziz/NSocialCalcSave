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

    public sealed class Cell
    {
        public object DataValue { get; set; }
        public CellDataType DataType { get; set; }
        public CellValueType ValueType { get; set; }
        public string Formula { get; set; }
        public bool ReadOnly { get; set; }
        public string Errors { get; set; }
        public int Bt { get; set; }
        public int Br { get; set; }
        public int Bb { get; set; }
        public int Bl { get; set; }
        public int Layout { get; set; }
        public int Font { get; set; }
        public int Color { get; set; }
        public int BgColor { get; set; }
        public int CellFormat { get; set; }
        public int NonTextValueFormat { get; set; }
        public int TextValueFormat { get; set; }
        public int ColSpan { get; set; }
        public int RowSpan { get; set; }
        public string Cssc { get; set; }
        public string Csss { get; set; }
        public string Comment { get; set; }
    }

    public sealed class NamedRange
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Definition { get; set; }
    }

    public sealed class Sheet
    {
        public ICollection<KeyValuePair<string, Cell>> Cells { get; set; }

        public ICollection<KeyValuePair<string, string>> ColWidths { get; set; }
        public ICollection<KeyValuePair<string, bool>> ColHides { get; set; }
        public ICollection<KeyValuePair<int, int>> RowHeights { get; set; }
        public ICollection<KeyValuePair<int, bool>> RowHides { get; set; }
        public ICollection<NamedRange> Names { get; set; }
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