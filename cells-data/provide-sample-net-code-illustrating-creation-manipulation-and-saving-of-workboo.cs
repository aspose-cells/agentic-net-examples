using System;
using Aspose.Cells;

namespace AsposeCellsSample
{
    public class WorkbookDemo
    {
        public static void Main()
        {
            // Create a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");

            // Add some sample data
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue(999.99);
            sheet.Cells["A3"].PutValue("Phone");
            sheet.Cells["B3"].PutValue(699.99);

            // Add a second worksheet for summary information
            int summaryIndex = workbook.Worksheets.Add();
            Worksheet summarySheet = workbook.Worksheets[summaryIndex];
            summarySheet.Name = "Summary";

            // Write summary data
            summarySheet.Cells["A1"].PutValue("Total Items");
            summarySheet.Cells["B1"].PutValue(2);

            // Save the workbook to an XLSX file
            workbook.Save("SampleWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}