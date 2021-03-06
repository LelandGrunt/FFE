﻿using ExcelDna.Integration.CustomUI;
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

        public void RecalculateAvqCallLimitReachedError(IRibbonControl control)
        {
            Avq.CallLimitReachedError = false;

            FfeExcel.FindCells((string)Avq.AvqExcelErrorCallLimitReached(), new ExcelFindOptions()
            {
                LookIn = XlFindLookIn.xlValues,
                LookAt = XlLookAt.xlWhole,
                MatchCase = true
            }).Refresh("QAV");
        }

        public void SetStopRefreshAtFirstCallLimitReachedError(IRibbonControl control, bool pressed)
        {
            AvqSetting.Default.StopRefreshAtFirstCallLimitReachedError = pressed;
            AvqSetting.Default.Save();
        }
    }
}