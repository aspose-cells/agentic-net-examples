// Title: Define a print area with repeating header rows and freeze those rows in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set print area and repeat header rows on each page | How to freeze top rows while defining print title rows in Aspose.Cells | C# example for using FreezePanes together with print titles in Aspose.Cells | Set print area A1:E22 and freeze first two rows using Aspose.Cells .NET | Aspose.Cells page setup print titles and freeze panes code sample
// Tags: print area setup Aspose.Cells | repeat title rows Aspose.Cells | freeze top rows Aspose.Cells | page setup API Aspose.Cells .NET | Excel workbook print configuration C#

using System;
using Aspose.Cells;

namespace PrintAreaAndFreezeDemo
{
    // The sample creates a workbook, adds header and data rows, defines the print area (A1:E22), sets rows 1‑2 as repeating title rows, freezes those rows in the UI with FreezePanes, auto‑fits columns, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Report";

                // Populate sample data (header + 20 rows of data)
                // Header rows (rows 1 and 2)
                sheet.Cells["A1"].PutValue("Company");
                sheet.Cells["B1"].PutValue("Quarter");
                sheet.Cells["C1"].PutValue("Revenue");
                sheet.Cells["D1"].PutValue("Expenses");
                sheet.Cells["E1"].PutValue("Profit");

                sheet.Cells["A2"].PutValue("Acme Corp");
                sheet.Cells["B2"].PutValue("Q1");
                sheet.Cells["C2"].PutValue(120000);
                sheet.Cells["D2"].PutValue(80000);
                sheet.Cells["E2"].PutValue(40000);

                // Data rows (rows 3‑22)
                for (int i = 3; i <= 22; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Company {i - 2}");
                    sheet.Cells[$"B{i}"].PutValue($"Q{((i - 3) % 4) + 1}");
                    sheet.Cells[$"C{i}"].PutValue(50000 + (i * 1000));
                    sheet.Cells[$"D{i}"].PutValue(30000 + (i * 800));
                    sheet.Cells[$"E{i}"].PutValue(20000 + (i * 200));
                }

                // Define the print area to include the header rows and all data rows (A1:E22)
                sheet.PageSetup.PrintArea = "A1:E22";

                // Set the rows that should repeat on each printed page (the header rows)
                sheet.PageSetup.PrintTitleRows = "1:2";

                // Freeze the first two rows (rows 1‑2) in the UI.
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Row/column are zero‑based split positions; freeze 2 rows, 0 columns.
                sheet.FreezePanes(2, 0, 2, 0);

                // Adjust column widths for better visibility
                sheet.AutoFitColumns();

                // Save the workbook
                string outputPath = "PrintAreaAndFreezeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("Print area set to: " + sheet.PageSetup.PrintArea);
                Console.WriteLine("Print title rows set to: " + sheet.PageSetup.PrintTitleRows);
                Console.WriteLine("Rows 1‑2 are frozen in the worksheet view.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
