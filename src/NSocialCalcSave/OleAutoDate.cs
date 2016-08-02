//
//   (C) 2001 Marcel Narings
// Copyright (C) 2004-2006 Novell, Inc (http://www.novell.com)
// Copyright (C) 2012 Xamarin Inc (http://www.xamarin.com)
//
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

namespace NSocialCalcSave
{
    using System;

    // Source: https://github.com/mono/mono/blob/98d56eb3038557261a6b8a5f66dc703a9a050609/mcs/class/corlib/System/DateTime.cs

    // ReSharper disable once InconsistentNaming

    public static class OleAutoDate
    {
        // for OLE Automation dates
        const long Ticks18991230 = 599264352000000000L;
        const double MinValue = -657435.0d;
        const double MaxValue = 2958466.0d;

        public static DateTime ToDateTime(double d)
        {
            // An OLE Automation date is implemented as a floating-point number
            // whose value is the number of days from midnight, 30 December 1899.

            // d must be negative 657435.0 through positive 2958466.0.
            if ((d <= MinValue) || (d >= MaxValue))
                throw new ArgumentException("Not a legal OleAut date.", nameof(d));

            var dt = new DateTime(Ticks18991230);
            if (d < 0.0d)
            {
                var days = Math.Ceiling(d);
                // integer part is the number of days (negative)
                dt = dt.AddRoundedMilliseconds(days * 86400000);
                // but decimals are the number of hours (in days fractions) and positive
                var hours = days - d;
                dt = dt.AddRoundedMilliseconds(hours * 86400000);
            }
            else
            {
                dt = dt.AddRoundedMilliseconds(d * 86400000);
            }

            return dt;
        }

        // required to match MS implementation for OADate (OLE Automation)

        static DateTime AddRoundedMilliseconds(this DateTime dateTime, double ms)
        {
            if (ms * TimeSpan.TicksPerMillisecond > long.MaxValue ||
                ms * TimeSpan.TicksPerMillisecond < long.MinValue)
            {
                throw new ArgumentOutOfRangeException();
            }
            var msticks = (long)(ms + (ms > 0 ? 0.5 : -0.5)) * TimeSpan.TicksPerMillisecond;
            return dateTime.AddTicks(msticks);
        }

        public static double From(DateTime dateTime)
        {
            var t = dateTime.Ticks;
            // uninitialized DateTime case
            if (t == 0)
                return 0;
            // we can't reach minimum value
            if (t < 31242239136000000)
                return MinValue + 0.001;

            var ts = new TimeSpan(dateTime.Ticks - Ticks18991230);
            var result = ts.TotalDays;
            // t < 0 (where 599264352000000000 == 0.0d for OA)
            if (t < 599264352000000000)
            {
                // negative days (int) but decimals are positive
                var d = Math.Ceiling(result);
                result = d - 2 - (result - d);
            }
            else
            {
                // we can't reach maximum value
                if (result >= MaxValue)
                    result = MaxValue - 0.00000001d;
            }
            return result;
        }
    }
}