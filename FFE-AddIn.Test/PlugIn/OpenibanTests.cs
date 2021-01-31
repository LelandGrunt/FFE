using Xunit;

namespace FFE
{
    public class OpenibanTests
    {
        [Trait("Version", "Release")]
        [Trait("Query", "QBIC")]
        [Theory]
        [InlineData("DE89370400440532013000")]
        public void QBicTest(string iban)
        {
            object actual = Openiban.QBIC(iban);
            Assert.Equal("COBADEFFXXX", actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "QCOUNTRIES")]
        [Fact]
        public void QCountriesTest()
        {
            object actual = Openiban.QCOUNTRIES();
            Assert.IsType<object[,]>(actual);
            Assert.NotEmpty((object[,])actual);
        }

        [Trait("Version", "Release")]
        [Trait("Query", "QIBAN")]
        [Theory]
        [InlineData("DE", "37040044", "0532013000")]
        public void QIbanTest(string countryCode, string bankCode, string accountNumber)
        {
            object actual = Openiban.QIBAN(countryCode, bankCode, accountNumber);
            Assert.Equal("DE89370400440532013000", actual);
        }
    }
}