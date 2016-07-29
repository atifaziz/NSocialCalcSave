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
    #region Imports

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Globalization;
    using Mannex;
    using Mannex.Collections.Generic;

    #endregion

    public class SocialCalcFormat
    {
        #region Sheet Save Format

        // linetype:param1:param2:...
        //
        // Linetypes are:
        //
        //    version:versionname - version of this format. Currently 1.5.
        //
        //    cell:coord:type:value...:type:value... - Types are as follows:
        //
        //       v:value - straight numeric value
        //       t:value - straight text/wiki-text in cell, encoded to handle \, :, newlines
        //       vt:fulltype:value - value with value type/subtype
        //       vtf:fulltype:value:formulatext - formula resulting in value with value type/subtype, value and text encoded
        //       vtc:fulltype:value:valuetext - formatted text constant resulting in value with value type/subtype, value and text encoded
        //       vf:fvalue:formulatext - formula resulting in value, value and text encoded (obsolete: only pre format version 1.1)
        //          fvalue - first char is "N" for numeric value, "T" for text value, "H" for HTML value, rest is the value
        //       e:errortext - Error text. Non-blank means formula parsing/calculation results in error.
        //       b:topborder#:rightborder#:bottomborder#:leftborder# - border# in sheet border list or blank if none
        //       l:layout# - number in cell layout list
        //       f:font# - number in sheet fonts list
        //       c:color# - sheet color list index for text
        //       bg:color# - sheet color list index for background color
        //       cf:format# - sheet cell format number for explicit format (align:left, etc.)
        //       cvf:valueformat# - sheet cell value format number (obsolete: only pre format v1.2)
        //       tvf:valueformat# - sheet cell text value format number
        //       ntvf:valueformat# - sheet cell non-text value format number
        //       colspan:numcols - number of columns spanned in merged cell
        //       rowspan:numrows - number of rows spanned in merged cell
        //       cssc:classname - name of CSS class to be used for cell when published instead of one calculated here
        //       csss:styletext - explicit CSS style information, encoded to handle :, etc.
        //       mod:allow - if "y" allow modification of cell for live "view" recalc
        //       comment:value - encoded text of comment for this cell (added in v1.5)
        //
        //    col:
        //       w:widthval - number, "auto" (no width in <col> tag), number%, or blank (use default)
        //       hide: - yes/no, no is assumed if missing
        //    row:
        //       hide - yes/no, no is assumed if missing
        //
        //    sheet:
        //       c:lastcol - number
        //       r:lastrow - number
        //       w:defaultcolwidth - number, "auto", number%, or blank (default->80)
        //       h:defaultrowheight - not used
        //       tf:format# - cell format number for sheet default for text values
        //       ntf:format# - cell format number for sheet default for non-text values (i.e., numbers)
        //       layout:layout# - default cell layout number in cell layout list
        //       font:font# - default font number in sheet font list
        //       vf:valueformat# - default number value format number in sheet valueformat list (obsolete: only pre format version 1.2)
        //       ntvf:valueformat# - default non-text (number) value format number in sheet valueformat list
        //       tvf:valueformat# - default text value format number in sheet valueformat list
        //       color:color# - default number for text color in sheet color list
        //       bgcolor:color# - default number for background color in sheet color list
        //       circularreferencecell:coord - cell coord with a circular reference
        //       recalc:value - on/off (on is default). If not "off", appropriate changes to the sheet cause a recalc
        //       needsrecalc:value - yes/no (no is default). If "yes", formula values are not up to date
        //       usermaxcol:value - maximum column to display, 0 for unlimited (default=0)
        //       usermaxrow:value - maximum row to display, 0 for unlimited (default=0)
        //
        //    name:name:description:value - name definition, name in uppercase, with value being "B5", "A1:B7", or "=formula";
        //                                  description and value are encoded.
        //    font:fontnum:value - text of font definition (style weight size family) for font fontnum
        //                         "*" for "style weight", size, or family, means use default (first look to sheet, then builtin)
        //    color:colornum:rgbvalue - text of color definition (e.g., rgb(255,255,255)) for color colornum
        //    border:bordernum:value - text of border definition (thickness style color) for border bordernum
        //    layout:layoutnum:value - text of vertical alignment and padding style for cell layout layoutnum (* for default):
        //                             vertical-alignment:vavalue;padding:topval rightval bottomval leftval;
        //    cellformat:cformatnum:value - text of cell alignment (left/center/right) for cellformat cformatnum
        //    valueformat:vformatnum:value - text of number format (see FormatValueForDisplay) for valueformat vformatnum (changed in v1.2)
        //    clipboardrange:upperleftcoord:bottomrightcoord - ignored -- from wikiCalc
        //    clipboard:coord:type:value:... - ignored -- from wikiCalc
        //
        // If this is clipboard contents, then there is also information to facilitate pasting:
        //
        //    copiedfrom:upperleftcoord:bottomrightcoord - range from which this was copied

        #endregion

        public static Sheet ParseSheetSave(string savedSheet) =>
            ParseSheetSave(savedSheet.SplitIntoLines());

        static Sheet ParseSheetSave(IEnumerable<string> savedSheetLines)
        {
            var cells = new List<KeyValuePair<string, Cell>>();

            var colWidths    = new List<KeyValuePair<string, string>>();
            var colHides     = new List<KeyValuePair<string, bool>>();
            var rowHeights   = new List<KeyValuePair<int, int>>();
            var rowHides     = new List<KeyValuePair<int, bool>>();
            var names        = new List<NamedRange>();
            var layouts      = new List<KeyValuePair<int, string>>();
            var fonts        = new List<KeyValuePair<int, string>>();
            var colors       = new List<KeyValuePair<int, string>>();
            var borderStyles = new List<KeyValuePair<int, string>>();
            var cellFormats  = new List<KeyValuePair<int, string>>();
            var valueFormats = new List<KeyValuePair<int, string>>();

            var lastCol                   = default(int);
            var lastRow                   = default(int);
            var defaultColWidth           = default(string);
            var defaultRowHeight          = default(int);
            var defaultTextFormat         = default(int);
            var defaultNonTextFormat      = default(int);
            var defaultLayout             = default(int);
            var defaultFont               = default(int);
            var defaultTextValueFormat    = default(int);
            var defaultNonTextValueFormat = default(int);
            var defaultColor              = default(int);
            var defaultBgColor            = default(int);
            var circularReferenceCell     = default(string);
            var recalc                    = default(string);
            var needsRecalc               = default(bool);
            var userMaxCol                = default(int);
            var userMaxRow                = default(int);

            var copiedFrom = default(string);

            foreach (var parts in from line in savedSheetLines select line.Split(':'))
            {
                var pe = parts.AsEnumerable().GetEnumerator();
                switch (pe.Read())
                {
                    case "cell":
                    {
                        var cell = pe.Read();
                        cells.Add(cell.AsKeyTo(ParseCell(pe)));
                        break;
                    }
                    case "col":
                    {
                        var coord = pe.Read();
                        for (var token = pe.TryRead(); token != null; token = pe.TryRead())
                        {
                            switch (token)
                            {
                                case "w": colWidths.Add(coord.AsKeyTo(pe.Read())); break; // must be text - could be auto or %, etc.
                                case "hide": colHides.Add(coord.AsKeyTo(IsYes(pe.Read()))); break;
                                default: throw new Exception($"Unknown col type item '{token}'");
                            }
                        }
                        break;
                    }
                    case "row":
                    {
                        var coord = ParseInt(pe.Read());
                        for (var t = pe.TryRead(); t != null; t = pe.TryRead())
                        {
                            switch (t)
                            {
                                case "h": rowHeights.Add(coord.AsKeyTo(ParseInt(pe.Read()))); break;
                                case "hide": rowHides.Add(coord.AsKeyTo(IsYes(pe.Read()))); break;
                                default:
                                    throw new Exception($"Unknown row type item '{t}'");
                            }
                        }
                        break;
                    }
                    case "sheet":
                    {
                        for (var token = pe.TryRead(); token != null; token = pe.TryRead())
                        {
                            switch (token)
                            {
                                case "c"          : lastCol = ParseInt(pe.Read()); break;
                                case "r"          : lastRow = ParseInt(pe.Read()); break;
                                case "w"          : defaultColWidth = pe.Read(); break;
                                case "h"          : defaultRowHeight = ParseInt(pe.Read()); break;
                                case "tf"         : defaultTextFormat = ParseInt(pe.Read()); break;
                                case "ntf"        : defaultNonTextFormat = ParseInt(pe.Read()); break;
                                case "layout"     : defaultLayout = ParseInt(pe.Read()); break;
                                case "font"       : defaultFont = ParseInt(pe.Read()); break;
                                case "tvf"        : defaultTextValueFormat = ParseInt(pe.Read()); break;
                                case "ntvf"       : defaultNonTextValueFormat = ParseInt(pe.Read()); break;
                                case "color"      : defaultColor = ParseInt(pe.Read()); break;
                                case "bgcolor"    : defaultBgColor = ParseInt(pe.Read()); break;
                                case "circularreferencecell": circularReferenceCell = pe.Read(); break;
                                case "recalc"     : recalc = pe.Read(); break;
                                case "needsrecalc": needsRecalc = IsYes(pe.Read()); break;
                                case "usermaxcol" : userMaxCol = ParseInt(pe.Read()); break;
                                case "usermaxrow" : userMaxRow = ParseInt(pe.Read()); break;
                                default: throw new Exception($"Unknown sheet attribute type item '{token}'");
                            }
                        }
                        break;
                    }
                    case "name"       : names.Add(new NamedRange
                                        {
                                            Name        = DecodeFromSave(pe.Read()).ToUpperInvariant(),
                                            Description = DecodeFromSave(pe.Read()),
                                            Definition  = DecodeFromSave(pe.Read())
                                        });
                                        break;
                    case "layout"     : layouts.Add(ParseInt(pe.Read()).AsKeyTo(string.Join(":", parts.Skip(2)) /* layouts can have ":" in them */)); break;
                    case "font"       : fonts.Add(ParseInt(pe.Read()).AsKeyTo(pe.Read())); break;
                    case "color"      : colors.Add(ParseInt(pe.Read()).AsKeyTo(pe.Read())); break;
                    case "border"     : borderStyles.Add(ParseInt(pe.Read()).AsKeyTo(pe.Read())); break;
                    case "cellformat" : cellFormats.Add(ParseInt(pe.Read()).AsKeyTo(DecodeFromSave(pe.Read()))); break;
                    case "valueformat": valueFormats.Add(ParseInt(pe.Read()).AsKeyTo(DecodeFromSave(pe.Read()))); break;
                    case "copiedfrom" : copiedFrom = pe.Read() + ":" + pe.Read(); break;

                    case "":
                    case "version":
                    case "clipboardrange": // in save versions up to 1.3. Ignored.
                    case "clipboard":
                        break;

                    default:
                        throw new Exception($"Unknown line type '{pe.Current}'");
                }
            }

            return new Sheet
            {
                Cells                     = cells,

                ColWidths                 = colWidths,
                ColHides                  = colHides,
                RowHeights                = rowHeights,
                RowHides                  = rowHides,
                Names                     = names,
                Layouts                   = layouts,
                Fonts                     = fonts,
                Colors                    = colors,
                BorderStyles              = borderStyles,
                CellFormats               = cellFormats,
                ValueFormats              = valueFormats,

                LastCol                   = lastCol,
                LastRow                   = lastRow,
                DefaultColWidth           = defaultColWidth,
                DefaultRowHeight          = defaultRowHeight,
                DefaultTextFormat         = defaultTextFormat,
                DefaultNonTextFormat      = defaultNonTextFormat,
                DefaultLayout             = defaultLayout,
                DefaultFont               = defaultFont,
                DefaultTextValueFormat    = defaultTextValueFormat,
                DefaultNonTextValueFormat = defaultNonTextValueFormat,
                DefaultColor              = defaultColor,
                DefaultBgColor            = defaultBgColor,
                CircularReferenceCell     = circularReferenceCell,
                Recalc                    = recalc,
                NeedsRecalc               = needsRecalc,
                UserMaxCol                = userMaxCol,
                UserMaxRow                = userMaxRow,

                CopiedFrom                = copiedFrom
            };
        }

        static string DecodeFromSave(string s) =>
            s.IndexOf('\\') < 0 ? s // for performace reasons: replace nothing takes up time
                                : s.Replace(@"\c", ":").Replace(@"\n", "\n").Replace(@"\b", @"\");

        //
        // SocialCalc.CellFromStringParts(sheet, cell, parts, j)
        //
        // Takes string that has been split by ":" in parts, starting at item j,
        // and fills in cell assuming save format.
        //

        static int ParseInt(string s) => int.Parse(s, CultureInfo.InvariantCulture);
        static double ParseNum(string s) => double.Parse(s, CultureInfo.InvariantCulture);

        static Cell ParseCell(IEnumerator<string> token)
        {
            var dataValue          = default(object);
            var dataType           = default(CellDataType);
            var valueType         = default(CellValueType);
            var formula            = default(string);
            var readOnly           = default(bool);
            var errors             = default(string);
            var bt                 = default(int);
            var br                 = default(int);
            var bb                 = default(int);
            var bl                 = default(int);
            var layout             = default(int);
            var font               = default(int);
            var color              = default(int);
            var bgcolor            = default(int);
            var cellFormat         = default(int);
            var nonTextValueFormat = default(int);
            var textValueFormat    = default(int);
            var colspan            = default(int);
            var rowspan            = default(int);
            var cssc               = default(string);
            var csss               = default(string);
            var comment            = default(string);

            while (token.MoveNext())
            {
                var type = token.Current;
                switch (type)
                {
                    // cell:coord:type:value...:type:value... - Types are as follows:
                    // v:value - straight numeric value
                    case "v":
                        dataValue = ParseNum(token.Read());
                        dataType = CellDataType.Number;
                        valueType = CellValueType.Number;
                        break;
                    // t:value - straight text/wiki-text in cell, encoded to handle \, :, newlines
                    case "t":
                        dataValue = DecodeFromSave(token.Read());
                        dataType = CellDataType.Text;
                        valueType = CellValueType.Text;
                        break;
                    // vt:fulltype:value - value with value type/subtype
                    case "vt":
                        if ((valueType = ParseCellValueType(token.Read())).IsNumeric())
                        {
                            dataType = CellDataType.Number;
                            dataValue = ParseNum(token.Read());
                        }
                        else
                        {
                            dataType = CellDataType.Text;
                            var v = DecodeFromSave(token.Read());
                            dataValue = valueType == CellValueType.Url ? (object)new Uri(v) : v;
                        }
                        break;
                    // vtf:fulltype:value:formulatext - formula resulting in value with value type/subtype, value and text encoded
                    // vtc:fulltype:value:valuetext - formatted text constant resulting in value with value type/subtype, value and text encoded
                    case "vtf":
                    case "vtc":
                        dataValue = (valueType = ParseCellValueType(token.Read())).IsNumeric()
                                  ? (valueType == CellValueType.Logical ? ParseInt(token.Read()) != 0 : (object)ParseNum(token.Read()))
                                  : DecodeFromSave(token.Read());
                        formula = DecodeFromSave(token.Read());
                        dataType = type[1] == 'c' ? CellDataType.Constant : CellDataType.Formula;
                        break;
                    case "ro": readOnly = IsYes(DecodeFromSave(token.Read())); break;
                    // e:errortext - Error text. Non-blank means formula parsing/calculation results in error.
                    case "e": errors = DecodeFromSave(token.Read()); break;
                    // b:topborder#:rightborder#:bottomborder#:leftborder# - border# in sheet border list or blank if none
                    case "b":
                        bt = ParseInt(token.Read());
                        br = ParseInt(token.Read());
                        bb = ParseInt(token.Read());
                        bl = ParseInt(token.Read());
                        break;
                    // l:layout# - number in cell layout list
                    case "l": layout = ParseInt(token.Read()); break;
                    // f:font# - number in sheet fonts list
                    case "f": font = ParseInt(token.Read()); break;
                    // c:color# - sheet color list index for text
                    case "c": color = ParseInt(token.Read()); break;
                    // bg:color# - sheet color list index for background color
                    case "bg": bgcolor = ParseInt(token.Read()); break;
                    // cf:format# - sheet cell format number for explicit format (align:left, etc.)
                    case "cf": cellFormat = ParseInt(token.Read()); break;
                    // ntvf:valueformat# - sheet cell non-text value format number
                    case "ntvf": nonTextValueFormat = ParseInt(token.Read()); break;
                    // tvf:valueformat# - sheet cell text value format number
                    case "tvf": textValueFormat = ParseInt(token.Read()); break;
                    // colspan:numcols - number of columns spanned in merged cell
                    case "colspan": colspan = ParseInt(token.Read()); break;
                    // rowspan:numrows - number of rows spanned in merged cell
                    case "rowspan": rowspan = ParseInt(token.Read()); break;
                    // cssc:classname - name of CSS class to be used for cell when published instead of one calculated here
                    case "cssc": cssc = token.Read(); break;
                    // csss:styletext - explicit CSS style information, encoded to handle :, etc.
                    case "csss": csss = DecodeFromSave(token.Read()); break;
                    // mod:allow - if "y" allow modification of cell for live "view" recalc
                    case "mod": token.Read(); break;
                    // comment:value - encoded text of comment for this cell (added in v1.5)
                    case "comment": comment = DecodeFromSave(token.Read()); break;
                    default:
                        // vf:fvalue:formulatext - formula resulting in value, value and text encoded (obsolete: only pre format version 1.1)
                        //    fvalue - first char is "N" for numeric value, "T" for text value, "H" for HTML value, rest is the value
                        // cvf:valueformat# - sheet cell value format number (obsolete: only pre format v1.2)
                        throw new Exception($"Unknown cell type item '{type}'");
                }
            }

            return new Cell
            {
                DataValue          = dataValue,
                DataType           = dataType,
                ValueType          = valueType,
                Formula            = formula,
                ReadOnly           = readOnly,
                Errors             = errors,
                Bt                 = bt,
                Br                 = br,
                Bb                 = bb,
                Bl                 = bl,
                Layout             = layout,
                Font               = font,
                Color              = color,
                BgColor            = bgcolor,
                CellFormat         = cellFormat,
                NonTextValueFormat = nonTextValueFormat,
                TextValueFormat    = textValueFormat,
                ColSpan            = colspan,
                RowSpan            = rowspan,
                Cssc               = cssc,
                Csss               = csss,
                Comment            = comment,
            };
        }

        static bool IsYes(string s) => "yes".Equals(s, StringComparison.OrdinalIgnoreCase);

        static CellValueType ParseCellValueType(string s)
        {
            if (s.Length == 0)
                return CellValueType.Undefined;

            switch (s[0])
            {
                case 'n':
                    if (s.Length == 1)
                        return CellValueType.Number;
                    switch (s[1])
                    {
                        case '%': return CellValueType.Percentage;
                        case '$': return CellValueType.Currency;
                        case 'l': return CellValueType.Logical;
                        case 'd':
                            switch (s.Length)
                            {
                                case 2: return CellValueType.Date;
                                case 3: if (s[2] == 't') return CellValueType.DateTime; else break;
                            }
                            break;
                    }
                    break;
                case 't':
                    switch (s.Length)
                    {
                        case 1: return CellValueType.Text;
                        case 2:
                            switch (s[1])
                            {
                                case 'h': return CellValueType.Html;
                                case 'l': return CellValueType.Url;
                            }
                            break;
                    }
                    break;
                case 'e':
                    switch (s)
                    {
                        case "e#N/A":
                        case "e#NULL!":
                        case "e#NUM!":
                        case "e#DIV/0!":
                        case "e#VALUE!":
                        case "e#REF!":
                        case "e#NAME?":
                            return CellValueType.Error;
                    }
                    break;
            }
            throw new Exception("Unknown cell value type: " + s);
        }
    }

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
        DateTime,   // ndt
        Error,      // e#...
    }

    static class CellValueTypeExtensions
    {
        public static bool IsNumeric(this CellValueType type) =>
                type == CellValueType.Number
            ||  type == CellValueType.Logical
            ||  type == CellValueType.Percentage
            ||  type == CellValueType.Currency
            ||  type == CellValueType.Date
            ||  type == CellValueType.DateTime;
    }
}
