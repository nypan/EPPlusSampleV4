using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OfficeOpenXml;

namespace EPPlusTestsamples
{
    [TestClass]
    public class EPPlusSampleTest
    {
        [TestMethod]

        [DeploymentItem("calc_amount.xlsx")]
        public void OpenXSLTSheetAndCalcFormula()
        {
            // Open Excel sheet file
            var xl = File.Open(@".\calc_amount.xlsx", FileMode.Open);
            // open excel file
            var pck = new ExcelPackage(xl);
            // Enter 
            // Quantity  
            var ws = pck.Workbook.Worksheets[1];
            var qnty = ws.Names["QUANTITY"];
            qnty.Value = 30;
            // Price per Quantity 
            var price = ws.Names["PRICE"];
            price.Value = 10;
            // OK 30 * 10 = 300;
            // Calc
            //ws.Calculate();
            pck.Workbook.Calculate();
            // Check Amount 
            var amount = ws.Names["AMOUNT"];
            var a2 = ws.Cells["C5"];
            Assert.AreEqual(300, amount.Value);
        }
    }
}
