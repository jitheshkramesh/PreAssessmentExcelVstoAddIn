using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace PreAssessmentExcelVstoAddIn
{
    public static class ExcelHelper
    {
        private static Dictionary<string, object> originalValues = new Dictionary<string, object>();

        public static void ConvertToAlphanumeric()
        {
            var app = Globals.ThisAddIn.Application;
            var range = app.Selection as Excel.Range;

            if (range == null) return;

            foreach (Excel.Range cell in range.Cells)
            {
                if (cell.Value2 != null)
                {
                    string address = cell.Address[false, false];
                    object original = cell.Value2;

                    if (!originalValues.ContainsKey(address))
                        originalValues[address] = original;

                    string cleaned = Regex.Replace(original.ToString(), "[^a-zA-Z0-9]", "");
                    cell.Value2 = cleaned;
                }
            }
        }

        public static void RevertOriginalValues()
        {
            var app = Globals.ThisAddIn.Application;
            var range = app.Selection as Excel.Range;

            if (range == null) return;

            foreach (Excel.Range cell in range.Cells)
            {
                string address = cell.Address[false, false];

                if (originalValues.ContainsKey(address))
                {
                    cell.Value2 = originalValues[address];
                    originalValues.Remove(address);
                }
            }
        }
    }
}
