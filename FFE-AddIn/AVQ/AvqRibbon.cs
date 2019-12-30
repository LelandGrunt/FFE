using ExcelDna.Integration.CustomUI;
using Microsoft.Office.Interop.Excel;

namespace FFE
{
    //[ComVisible(true)]
    public partial class FfeRibbon : ExcelRibbon
    {
        private const string AVQ_NAME = "AVQ";

        public void SetAvApiKey(IRibbonControl control)
        {
            string apiKey = null;

            apiKey = Microsoft.VisualBasic.Interaction.InputBox("Alpha Vantage API Key:", AVQ_NAME, string.IsNullOrEmpty(Avq.ApiKey) ? "<Enter your Alpha Vantage API Key here>" : Avq.ApiKey);
            if (!string.IsNullOrEmpty(apiKey))
            {
                Avq.ApiKey = apiKey;
                excelApp.StatusBar = null;
            }
        }

        public void SetBatchQueryMode(IRibbonControl control, string id, int index)
        {
            AvqSetting.Default.BatchQueryMode = id;
            AvqSetting.Default.Save();

            /*if (id.Equals("ddcAvqBatchQueryModeItemNamedRange"))
            {
                // Enable ebxAvqRangeName Editbox.
            }
            else
            {
                // Disable ebxAvqRangeName Editbox.
            }*/
        }

        public void SetBatchQueryNamedRange(IRibbonControl control, string value)
        {
            AvqSetting.Default.BatchQueryNamedRange = value;
            AvqSetting.Default.Save();
        }

        public void QueryAlphaVantageAsBatch(IRibbonControl control) { Avq.QueryAlphaVantageAsBatch(); }

        public void RecalculateAvqCallLimitReachedError(IRibbonControl control)
        {
            FfeExcel.FindCells((string)Avq.AvqExcelErrorCallLimitReached(), new ExcelFindOptions()
            {
                LookIn = XlFindLookIn.xlValues,
                LookAt = XlLookAt.xlWhole,
                MatchCase = true
            }).Recalculate();
            //IDEA: Stop recalculation if current calculation returns an AvqExcelErrorCallLimitReached error?
        }
    }
}