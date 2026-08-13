// Title: C# Aspose.Cells Example: Hide Excel Table Header Row While Keeping Data Rows Visible
// Description: A concise C# sample that creates a workbook, adds a ListObject (Excel table), applies a style, and hides the table header by setting ShowHeaderRow to false, leaving the data rows visible for reporting or export.
// Keywords: Aspose.Cells | C# | .NET | Hide Excel table header | ListObject ShowHeaderRow | Excel table styling | Workbook export | Reporting without header | Sample code | GitHub example
// Common Searches: Aspose.Cells hide table header C# | ShowHeaderRow false example | C# hide Excel table header row | ListObject hide header Aspose.Cells | How to hide table header in Aspose.Cells
// Developer Intent: Hide the header row of an Excel ListObject while preserving all data rows using Aspose.Cells for .NET.
// Use Cases: Generate printable reports where column titles are supplied externally. | Create data‑only export files that will be merged with a custom header later. | Prepare workbooks for downstream processing that requires raw rows without a table header.
// AI Prompts: Write C# code with Aspose.Cells that adds a ListObject to a worksheet and hides its header row using ShowHeaderRow = false. | Explain the effect of the ShowHeaderRow property on Excel table rendering in Aspose.Cells and how to toggle it for multiple tables. | Provide a GitHub‑ready example that hides an Excel table header while keeping the style and saves the workbook to a specified path.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // A concise C# sample that creates a workbook, adds a ListObject (Excel table), applies a style, and hides the table header by setting ShowHeaderRow to false, leaving the data rows visible for reporting or export.
    public class HideTableHeaderRow
    {
        public static void Main(string[] args)
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
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (including a header row)
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Banana");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("Cherry");
            worksheet.Cells["B4"].PutValue(30);

            // Add a ListObject (table) that includes the header row
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 2, true);
            ListObject table = worksheet.ListObjects[tableIndex];

            // Apply a style (optional)
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Hide the header row of the table while keeping data rows visible
            table.ShowHeaderRow = false;

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HideTableHeaderRow.xlsx");

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
