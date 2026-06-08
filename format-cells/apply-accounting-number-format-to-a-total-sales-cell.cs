using System;
using Aspose.Cells;

namespace AsposeCellsAccountingFormatDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value representing total sales into cell B2
            Cell totalSalesCell = sheet.Cells["B2"];
            totalSalesCell.PutValue(123456.78);

            // Create a style and set its number format to Accounting (value 37)
            Style accountingStyle = workbook.CreateStyle();
            accountingStyle.Number = 37; // Accounting format "#,##0_);(#,##0)"

            // Apply the style to the total sales cell
            totalSalesCell.SetStyle(accountingStyle);

            // Save the workbook
            workbook.Save("TotalSales_AccountingFormat.xlsx");
        }
    }
}