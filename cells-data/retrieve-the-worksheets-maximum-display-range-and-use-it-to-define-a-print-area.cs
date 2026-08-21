// Title: C# Example: Set Worksheet Print Area from MaxDisplayRange with Aspose.Cells
// Description: Demonstrates how to create a workbook, add data, merged cells and a shape, retrieve the worksheet's MaxDisplayRange, convert its bounds to A1 notation, assign the range to PageSetup.PrintArea, and save the file. The sample shows a reliable way to define a print area that automatically covers all visible content.
// Keywords: Aspose.Cells C# MaxDisplayRange | set print area programmatically | PageSetup.PrintArea Aspose | worksheet MaxDisplayRange example | include merged cells and shapes in print range | Aspose.Cells .NET tutorial | dynamic print area calculation | CellsHelper.CellIndexToName usage | GitHub Aspose.Cells demo | coding‑agent snippet
// Common Searches: Aspose.Cells set print area from MaxDisplayRange C# | how to include shapes in print range Aspose.Cells | retrieve maximum display range worksheet Aspose | C# code to define PageSetup.PrintArea automatically | MaxDisplayRange vs UsedRange Aspose.Cells
// Developer Intent: Obtain the worksheet’s maximum display range and apply it as the print area in a .NET workbook.
// Use Cases: Automatically configure the print area for reports that contain data, merged regions, and graphics. | Create a reusable helper method that sets PageSetup.PrintArea based on MaxDisplayRange for any worksheet. | Generate print‑ready Excel files where the printable region adapts to content changes without manual adjustments.
// AI Prompts: Write a C# method that takes a Worksheet object, gets its MaxDisplayRange, builds an A1‑style address, and sets PageSetup.PrintArea, handling null ranges. | Show how to use CellsHelper.CellIndexToName with MaxDisplayRange properties to construct the print area string. | Explain why MaxDisplayRange is preferred over UsedRange when the sheet contains shapes or merged cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For Shape-related classes

namespace MaxDisplayRangePrintAreaDemo
{
    // Demonstrates how to create a workbook, add data, merged cells and a shape, retrieve the worksheet's MaxDisplayRange, convert its bounds to A1 notation, assign the range to PageSetup.PrintArea, and save the file. The sample shows a reliable way to define a print area that automatically covers all visible content.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate the worksheet with sample data
                worksheet.Cells["A1"].PutValue("Header1");
                worksheet.Cells["B1"].PutValue("Header2");
                worksheet.Cells["A2"].PutValue(100);
                worksheet.Cells["B2"].PutValue(200);
                worksheet.Cells["A3"].PutValue(300);
                worksheet.Cells["B3"].PutValue(400);

                // Add a merged cell to demonstrate MaxDisplayRange includes it
                worksheet.Cells.Merge(4, 0, 2, 2); // Merge cells A5:C6

                // Add a text effect shape (required parameters include height and width)
                worksheet.Shapes.AddTextEffect(
                    MsoPresetTextEffect.TextEffect1, // preset effect
                    "Merged Area",                  // text
                    "Arial",                        // font name
                    12,                             // font size
                    false,                          // bold
                    false,                          // italic
                    4, 0,                           // upper‑left row & column
                    0, 0,                           // top & left offsets
                    0, 0);                          // height & width (auto‑size)

                // Retrieve the maximum display range (includes data, merged cells, shapes)
                var maxDisplayRange = worksheet.Cells.MaxDisplayRange;

                if (maxDisplayRange != null)
                {
                    // Calculate the start and end cell addresses of the range
                    string startCell = CellsHelper.CellIndexToName(maxDisplayRange.FirstRow, maxDisplayRange.FirstColumn);
                    string endCell = CellsHelper.CellIndexToName(
                        maxDisplayRange.FirstRow + maxDisplayRange.RowCount - 1,
                        maxDisplayRange.FirstColumn + maxDisplayRange.ColumnCount - 1);

                    // Define the print area using the calculated range
                    worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

                    Console.WriteLine($"Print area set to: {worksheet.PageSetup.PrintArea}");
                }
                else
                {
                    Console.WriteLine("Worksheet is empty; no print area defined.");
                }

                // Save the workbook
                string outputPath = "MaxDisplayRangePrintAreaDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
