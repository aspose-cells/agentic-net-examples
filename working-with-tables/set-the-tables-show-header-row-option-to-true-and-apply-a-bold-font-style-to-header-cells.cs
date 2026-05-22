using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class TableHeaderBoldDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate header row
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");

                // Populate some data rows
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(2.5);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(1.8);
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(3.2);

                // Define the range of the table (including header row)
                int firstRow = 0;      // zero‑based index for row 1
                int firstColumn = 0;   // column A
                int lastRow = 4;       // row 5 (zero‑based)
                int lastColumn = 1;    // column B

                // Add a ListObject (table) to the worksheet
                int tableIndex = worksheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowHeaderRow = true;

                // Create a style with bold font for the header row
                Style boldStyle = workbook.CreateStyle();
                boldStyle.Font.IsBold = true;
                StyleFlag flag = new StyleFlag { FontBold = true };
                worksheet.Cells.ApplyRowStyle(firstRow, boldStyle, flag);

                // Save the workbook
                string outputPath = "TableHeaderBoldDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            TableHeaderBoldDemo.Run();
        }
    }
}