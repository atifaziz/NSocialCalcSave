#region (C) 2018 Atif Aziz
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
    using System.Collections.Generic;

    struct StringSpan
    {
        readonly string _str;
        readonly int _start;
        readonly int _end;

        public StringSpan(string str) :
            this(str, 0) {}

        public StringSpan(string str, int start) :
            this(str, start, str.Length) {}

        public StringSpan(string str, int start, int end)
        {
            _str = str;
            _start = start;
            _end = end;
        }

        public char this[int i] => _str[_start + i];
        public int Length => _end - _start;

        //public StringCursor? IndexOf(char ch) =>
        //    _str.IndexOf(ch, _start) is int i && i >= 0 ? this + (i - _start) : (StringCursor?) null;

        public IEnumerable<StringSpan> Split(char separator)
        {
            for (var start = _start;;)
            {
                var i = _str.IndexOf(separator, start);
                if (i < 0 || i >= _end)
                {
                    yield return new StringSpan(_str, start, _end);
                    break;
                }
                yield return new StringSpan(_str, start, i);
                start = i + 1;
            }
        }

        public IEnumerable<StringSpan> SplitIntoLines()
        {
            // ReSharper disable once LoopCanBeConvertedToQuery
            foreach (var line in Split('\n'))
                yield return line.TrimEnd('\r');
        }

        public StringSpan Trim(char ch) =>
            TrimStart(ch).TrimEnd(ch);

        public StringSpan TrimEnd(char ch)
        {
            var end = _end;
            while (end > _start && _str[end - 1] == ch)
                end--;
            return new StringSpan(_str, _start, end);
        }

        public StringSpan TrimStart(char ch)
        {
            var start = _start;
            while (start < _end && _str[start] == ch)
                start++;
            return new StringSpan(_str, start, _end);
        }

        public override string ToString() =>
            _str?.Substring(_start, Length) ?? string.Empty;

        public static StringSpan operator++(StringSpan str) => str + 1;
        public static StringSpan operator--(StringSpan str) => str - 1;

        public static StringSpan operator+(StringSpan str, int i) =>
            new StringSpan(str._str, str._start + i, str._end);

        public static StringSpan operator-(StringSpan str, int i) =>
            new StringSpan(str._str, str._start - i, str._end);

        public static bool operator !=(StringSpan a, StringSpan b) => !(a == b);

        public static bool operator==(StringSpan a, StringSpan b)
            => a.Length == b.Length
            && 0 == string.Compare(a._str, a._start, b._str, b._start, a.Length, StringComparison.Ordinal);

        public static implicit operator StringSpan(string str) =>
            new StringSpan(str);

        [Obsolete]
        public static implicit operator string(StringSpan str) =>
            str.ToString();
    }
}
