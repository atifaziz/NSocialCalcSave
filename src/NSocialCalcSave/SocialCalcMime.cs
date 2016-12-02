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
    using System;
    using System.Linq;
    using System.Text.RegularExpressions;
    using Mannex;

    public static class SocialCalcMime
    {
        /// <remarks>
        /// This method is currently designed to work with a very specific
        /// version of the format where the multipart/mixed content uses a
        /// boundary of <c>SocialCalcSpreadsheetControlSave</c> and the sheet
        /// is the second part.
        /// </remarks>

        public static ISheet ParseSheet(string source) =>
            SocialCalcFormat.ParseSheetSave(ParseSheetSource(source));

        /// <remarks>
        /// This method is currently designed to work with a very specific
        /// version of the format where the multipart/mixed content uses a
        /// boundary of <c>SocialCalcSpreadsheetControlSave</c> and the sheet
        /// is the second part.
        /// </remarks>

        public static string ParseSheetSource(string source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            //
            // The source will look something like this:
            //
            //     socialcalc:version:1.0
            //     MIME-Version: 1.0
            //     Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave
            //     --SocialCalcSpreadsheetControlSave
            //     Content-type: text/plain; charset=UTF-8
            //
            //     # SocialCalc Spreadsheet Control Save
            //     version:1.0
            //     part:sheet
            //     part:edit
            //     part:audit
            //     --SocialCalcSpreadsheetControlSave
            //     Content-type: text/plain; charset=UTF-8
            //
            //     version:1.5
            //     ...
            //     --SocialCalcSpreadsheetControlSave
            //     Content-type: text/plain; charset=UTF-8
            //
            //     version:1.0
            //     rowpane:0:1:1
            //     colpane:0:1:1
            //     ecell:B1
            //     --SocialCalcSpreadsheetControlSave
            //     Content-type: text/plain; charset=UTF-8
            //
            //     --SocialCalcSpreadsheetControlSave--
            //
            // The interesting bit is in part two. Short of bringing in a full
            // multi-part MIME parser (which may still be done later), cheat
            // and assume the source is formatted like above.
            //

            const string boundaryPattern = "--SocialCalcSpreadsheetControlSave(--)?\r?\n";

            var parts = Regex.Split(source, boundaryPattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            if (parts.Length < 3)
                throw new FormatException("Invalid SocialCalc MIME format.");

            // Get the body of the sheet.

            return string.Join(Environment.NewLine,
                parts[2]
                    .SplitIntoLines()
                    .SkipWhile(s => s.Length > 0) // skip headers
                    .Skip(1));                    // skip blank line after headers
        }
    }
}