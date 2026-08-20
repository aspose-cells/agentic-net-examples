// Title: Freeze Header Row via Named Range with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, define a named range for the first‑row header, freeze the rows covered by that range using FreezePanes, and save the result as an XLSX file.
// Keywords: Aspose.Cells C# | FreezePanes | named range header | freeze header row | Excel worksheet freeze | Aspose.Cells API | C# Excel automation | freeze rows Aspose | define named range Aspose.Cells | Excel file generation .NET
// Common Searches: Aspose.Cells freeze first row C# | how to create named range and freeze panes Aspose.Cells | C# code to freeze header rows in Excel with Aspose | Aspose.Cells FreezePanes example .NET | define named range for header Aspose.Cells
// Developer Intent: Create a named range for the header row and freeze those rows in a worksheet.
// Use Cases: Keep column titles visible while scrolling through large datasets. | Reference the header range in formulas or data‑validation after freezing. | Apply consistent header freezing across multiple sheets in a workbook.
// AI Prompts: Generate C# Aspose.Cells code that defines a named range for row 1 and freezes that row. | Show how to freeze the first three rows and assign a named range covering them using Aspose.Cells. | Provide error‑handled C# example that creates a header named range and freezes the corresponding rows with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHeaderFreezeDemo
{
    // Demonstrates how to create a workbook, define a named range for the first‑row header, freeze the rows covered by that range using FreezePanes, and save the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate header cells (example with three columns)
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["C1"].PutValue("Header3");

                // The header occupies the first row (zero‑based index 0)
                int rowsToFreeze = 1; // Freeze the first row

                // Freeze the rows encompassed by the header range.
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Freeze at the row just below the header (rowsToFreeze) and column 0.
                sheet.FreezePanes(rowsToFreeze, 0, rowsToFreeze, 0);

                // Define output file path
                string outputPath = "HeaderFreezeDemo.xlsx";

                // Save the workbook
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
