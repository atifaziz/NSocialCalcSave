namespace NSocialCalcSave.Tests
{
    using System.Linq;
    using NUnit.Framework;

    public class StringSpanTests
    {
        [Test]
        public void DefaultLengthIsZero()
        {
            StringSpan @default = default;
            Assert.That(@default.Length, Is.EqualTo(0));
        }

        [Test]
        public void DefaultStringIsEmpty()
        {
            StringSpan @default = default;
            Assert.That(@default.ToString(), Is.EqualTo(string.Empty));
        }

        [TestCase("")]
        [TestCase(",")]
        [TestCase("foo")]
        [TestCase("foo,bar")]
        [TestCase("foo,bar,baz")]
        public void Split(string s)
        {
            var ss = new StringSpan(s);
            var parts =
                from p in ss.Split(',')
                select p.ToString();
            Assert.That(parts, Is.EqualTo(s.Split(',')));
        }

        [TestCase("", '-', "")]
        [TestCase("-", '-', "")]
        [TestCase("--", '-', "")]
        [TestCase("---", '-', "")]
        [TestCase("-foo-", '-', "foo-")]
        [TestCase("--foo--", '-', "foo--")]
        [TestCase("---foo---", '-', "foo---")]
        public void TrimStart(string s, char ch, string expected)
        {
            var ss = new StringSpan(s);
            ss = ss.TrimStart(ch);
            Assert.That(ss.ToString(), Is.EqualTo(expected));
        }

        [TestCase("", '-', "")]
        [TestCase("-", '-', "")]
        [TestCase("--", '-', "")]
        [TestCase("---", '-', "")]
        [TestCase("-foo-", '-', "-foo")]
        [TestCase("--foo--", '-', "--foo")]
        [TestCase("---foo---", '-', "---foo")]
        public void TrimEnd(string s, char ch, string expected)
        {
            var ss = new StringSpan(s);
            ss = ss.TrimEnd(ch);
            Assert.That(ss.ToString(), Is.EqualTo(expected));
        }

        [Test]
        public void SplitIntoLines()
        {
            var ss = new StringSpan("\nfoo\n\nbar\r\n\r\r\nbaz\r\r\nqux\n");
            var lines = ss.SplitIntoLines();
            Assert.That(lines, Is.EqualTo(new StringSpan[]
            {
                string.Empty,
                "foo",
                string.Empty,
                "bar",
                string.Empty,
                "baz",
                "qux",
                string.Empty,
            }));
        }
    }
}
