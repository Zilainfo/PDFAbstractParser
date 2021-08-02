using ConsoleApp;
using NUnit.Framework;

namespace ConsoleAppTest
{
    public class ParserTests
    {
        private TestContext _fixtureContext;


        [SetUp]
        public void Setup()
        {
            _fixtureContext = TestContext.CurrentContext;
        }

        [Test]
        public void ParserConstructor_CreatingTest()
        {
            NUnitUserInput input = new NUnitUserInput();
            input.fromPath = @"C:\Users\kstas\Desktop\Work\Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf";
            input.InPath = @"C:\Users\kstas\Desktop\Work\Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format.xlsx";

            Parser parser = new Parser(input, 42);

            Assert.That(parser.fromPath, Is.EqualTo(@"C:\Users\kstas\Desktop\Work\Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf"));
            Assert.That(parser.InPath, Is.EqualTo(@"C:\Users\kstas\Desktop\Work\Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format.xlsx"));
        }
    }
}
