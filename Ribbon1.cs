using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PreAssessmentExcelVstoAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelper.ConvertToAlphanumeric();
        }

        private void btnRevert_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelper.RevertOriginalValues();
        }
    }
}
