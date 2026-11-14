// Title: Generate a summary worksheet that aggregates amounts from multiple ListObject tables using structured reference formulas in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# sum column values from a ListObject table using structured references | How to create a summary sheet that totals amounts from two tables in a workbook with Aspose.Cells | C# code to add a summary worksheet with currency formatting for aggregated table totals in Aspose.Cells | Using Aspose.Cells to calculate combined total from Table1 and Table2 with =SUM(Table[Column]) | Programmatically apply currency number format to formula cells in Aspose.Cells .NET
// Tags: structured reference aggregation Aspose.Cells | sum ListObject column Aspose.Cells C# | summary worksheet generation Aspose.Cells | currency number format Aspose.Cells | multiple table totals Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsSummaryExample
{
    // The example creates a new workbook, adds two worksheets each containing a ListObject table (Table1 and Table2) with item and amount data, then inserts a Summary worksheet that uses structured reference formulas (e.g., =SUM(Table1[Amount])) to compute individual and combined totals. The result cells are formatted as currency and the workbook is saved as 'SummaryWorkbook.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Create first data worksheet with a table (Table1)
                // -------------------------------------------------
                Worksheet dataSheet1 = workbook.Worksheets[0];
                dataSheet1.Name = "Data1";

                // Populate headers
                dataSheet1.Cells["A1"].PutValue("Item");
                dataSheet1.Cells["B1"].PutValue("Amount");

                // Populate sample data
                dataSheet1.Cells["A2"].PutValue("Apple");
                dataSheet1.Cells["B2"].PutValue(120);
                dataSheet1.Cells["A3"].PutValue("Banana");
                dataSheet1.Cells["B3"].PutValue(80);
                dataSheet1.Cells["A4"].PutValue("Cherry");
                dataSheet1.Cells["B4"].PutValue(150);

                // Define the range that will become a table (including headers)
                int firstRow = 0;          // zero‑based index for row 1
                int firstColumn = 0;       // column A
                int totalRows = 4;         // rows 1‑4
                int totalColumns = 2;      // columns A‑B

                // Add ListObject (table) to the range
                int tableIndex1 = dataSheet1.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
                ListObject table1 = dataSheet1.ListObjects[tableIndex1];
                table1.DisplayName = "Table1";                     // set table name
                table1.ShowHeaderRow = true;
                table1.ShowTableStyleColumnStripes = true;
                table1.TableStyleType = TableStyleType.TableStyleMedium9;

                // -------------------------------------------------
                // Create second data worksheet with a table (Table2)
                // -------------------------------------------------
                Worksheet dataSheet2 = workbook.Worksheets.Add("Data2");

                // Populate headers
                dataSheet2.Cells["A1"].PutValue("Product");
                dataSheet2.Cells["B1"].PutValue("Amount");

                // Populate sample data
                dataSheet2.Cells["A2"].PutValue("Desk");
                dataSheet2.Cells["B2"].PutValue(300);
                dataSheet2.Cells["A3"].PutValue("Chair");
                dataSheet2.Cells["B3"].PutValue(200);
                dataSheet2.Cells["A4"].PutValue("Lamp");
                dataSheet2.Cells["B4"].PutValue(90);

                // Define the range for the second table
                firstRow = 0;
                firstColumn = 0;
                totalRows = 4;
                totalColumns = 2;

                // Add ListObject (table) to the range
                int tableIndex2 = dataSheet2.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
                ListObject table2 = dataSheet2.ListObjects[tableIndex2];
                table2.DisplayName = "Table2";                     // set table name
                table2.ShowHeaderRow = true;
                table2.ShowTableStyleColumnStripes = true;
                table2.TableStyleType = TableStyleType.TableStyleMedium14;

                // -------------------------------------------------
                // Create summary worksheet with structured reference formulas
                // -------------------------------------------------
                Worksheet summarySheet = workbook.Worksheets.Add("Summary");

                // Labels
                summarySheet.Cells["A1"].PutValue("Total Amount from Table1");
                summarySheet.Cells["A2"].PutValue("Total Amount from Table2");
                summarySheet.Cells["A3"].PutValue("Combined Total Amount");

                // Structured reference formulas
                summarySheet.Cells["B1"].Formula = "=SUM(Table1[Amount])";
                summarySheet.Cells["B2"].Formula = "=SUM(Table2[Amount])";
                summarySheet.Cells["B3"].Formula = "=SUM(Table1[Amount]) + SUM(Table2[Amount])";

                // Format the result cells as currency
                Style currencyStyle = workbook.CreateStyle();
                currencyStyle.Number = 164; // Currency format
                StyleFlag flag = new StyleFlag { NumberFormat = true };
                summarySheet.Cells["B1"].SetStyle(currencyStyle, flag);
                summarySheet.Cells["B2"].SetStyle(currencyStyle, flag);
                summarySheet.Cells["B3"].SetStyle(currencyStyle, flag);

                // -------------------------------------------------
                // Save the workbook to a file
                // -------------------------------------------------
                string outputPath = "SummaryWorkbook.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
