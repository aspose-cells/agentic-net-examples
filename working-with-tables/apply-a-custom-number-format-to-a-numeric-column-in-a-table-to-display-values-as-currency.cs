using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class TableCurrencyFormatDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1234.5);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(5678.9);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(9012.34);

            // Create a table (ListObject) that includes the data range
            // Parameters: first row, first column, last row, last column, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            // Optional: set a display name for the table (if supported)
            // table.DisplayName = "SalesTable";

            // Define a custom currency number format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Custom = "$#,##0.00";

            // Prepare a StyleFlag to apply only the number format
            StyleFlag flag = new StyleFlag { NumberFormat = true };

            // Determine the range of the numeric column (second column of the table)
            // DataRange gives the body of the table without the header row
            Aspose.Cells.Range dataRange = table.DataRange;
            int firstDataRow = dataRange.FirstRow;                     // first row of data (excluding header)
            int amountColumnIndex = dataRange.FirstColumn + 1;        // second column (Amount)
            int rowCount = dataRange.RowCount;                        // number of data rows

            // Create a range that covers the entire Amount column within the table
            Aspose.Cells.Range amountRange = sheet.Cells.CreateRange(firstDataRow, amountColumnIndex, rowCount, 1);

            // Apply the custom currency format to the range
            amountRange.ApplyStyle(currencyStyle, flag);

            // Save the workbook
            string outputPath = "TableCurrencyFormat.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
    }
}