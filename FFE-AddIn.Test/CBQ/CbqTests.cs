using Xunit;

namespace FFE
{
    public class CbqTests
    {
        [Trait("Query", "CBQ")]
        [Theory]
        [InlineData("870747", null, null)]
        [InlineData("870747", null, "price")]
        [InlineData("870747", null, "Price")]
        [InlineData("870747", null, "PRICE")]
        [InlineData("US5949181045", null, "Price")]
        [InlineData("870747", "GAT", null)]
        [InlineData("870747", "GER", "price")]
        [InlineData("US5949181045", "GAT", "Price")]
        public void QCbTest(string wrk_isin, string exchange, string info)
        {
            object actual = Cbq.QCB(wrk_isin, exchange: exchange, info: info);
            Assert.IsType<decimal>(actual);
        }

        /* Not testable due to missing ExcelReference.
        [Trait("Query", "CBQ")]
        [Theory]
        [InlineData("870747", "GAT", null)]
        public void QCbFTest(string wrk_isin, string exchange, string info)
        {
            object actual = Cbq.QCBF(wrk_isin, exchange: exchange, info: info);
            Assert.IsType<decimal>(actual);
        }*/
    }
}