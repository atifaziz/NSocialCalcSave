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
    public enum CellDataType
    {
        Undefined,
        Text,       // t
        Number,     // n/v
        Formula,    // f
        Constant,   // c = constant that is not a simple number (like "$1.20")
    }

    public enum CellValueType
    {
        Undefined,
        Text,       // t
        Html,       // th
        Url,        // tl
        Number,     // n
        Logical,    // nl
        Percentage, // n%
        Currency,   // n$
        Date,       // nd
        Time,       // nt
        DateTime,   // ndt
        // ReSharper disable once InconsistentNaming
        ErrorNA,    // e#N/A
        ErrorNull,  // e#NULL!
        ErrorNum,   // e#NUM!
        ErrorDiv0,  // e#DIV/0!
        ErrorValue, // e#VALUE!
        ErrorRef,   // e#REF!
        ErrorName,  // e#NAME?
    }

    public interface ICell
    {
        object DataValue { get; }
        CellDataType DataType { get; }
        CellValueType ValueType { get; }
        string Formula { get; }
        bool ReadOnly { get; }
        string Errors { get; }
        int Bt { get; }
        int Br { get; }
        int Bb { get; }
        int Bl { get; }
        int Layout { get; }
        int Font { get; }
        int Color { get; }
        int BgColor { get; }
        int CellFormat { get; }
        int NonTextValueFormat { get; }
        int TextValueFormat { get; }
        int ColSpan { get; }
        int RowSpan { get; }
        string Cssc { get; }
        string Csss { get; }
        string Comment { get; }
    }

    public sealed class Cell : ICell
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
}