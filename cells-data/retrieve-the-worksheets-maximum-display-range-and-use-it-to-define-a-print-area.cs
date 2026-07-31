// Title: Set Excel Print Area from MaxDisplayRange using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate data (including merged cells), retrieve the worksheet's MaxDisplayRange, convert its boundaries to A1 notation, assign the range to PageSetup.PrintArea, and save the file while safely handling a null range.
// Keywords: Aspose.Cells | C# | MaxDisplayRange | PrintArea | PageSetup | merged cells | Excel printing | dynamic print area | worksheet display range
// Common Searches: Aspose.Cells set print area from MaxDisplayRange C# | how to use MaxDisplayRange for printing in Aspose.Cells | define Excel print area programmatically .NET | include merged cells in Aspose.Cells print area | retrieve worksheet display range for print setup
// Developer Intent: Determine the worksheet’s maximum display range and apply it as the print area in an Aspose.Cells workbook.
// Use Cases: Automatically size the print area for reports that generate data at runtime. | Create a reusable helper that sets PageSetup.PrintArea to the content region of each worksheet. | Prevent blank pages by printing only the area that contains data, merged cells, or shapes.
// AI Prompts: Generate a C# method that takes a Worksheet, gets its MaxDisplayRange, and sets PageSetup.PrintArea, handling null ranges. | Write code to loop through all worksheets in a Workbook and assign each one a print area based on its MaxDisplayRange using Aspose.Cells. | Explain the differences between MaxDisplayRange and UsedRange in Aspose.Cells and advise when to use each for defining print areas.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    // Demonstrates how to create a workbook, populate data (including merged cells), retrieve the worksheet's MaxDisplayRange, convert its boundaries to A1 notation, assign the range to PageSetup.PrintArea, and save the file while safely handling a null range.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate some data to generate a display range
                worksheet.Cells["A1"].PutValue("Header1");
                worksheet.Cells["B1"].PutValue("Header2");
                worksheet.Cells["A2"].PutValue(100);
                worksheet.Cells["B2"].PutValue(200);
                worksheet.Cells["A3"].PutValue(300);
                worksheet.Cells["B3"].PutValue(400);

                // Add a merged cell to ensure it is included in MaxDisplayRange
                worksheet.Cells.Merge(4, 0, 1, 2); // Merge cells A5:C5
                worksheet.Cells["A5"].PutValue("Merged Cell");

                // Retrieve the maximum display range (includes data, merged cells, shapes)
                Aspose.Cells.Range maxDisplayRange = worksheet.Cells.MaxDisplayRange;

                // Guard against empty worksheet (MaxDisplayRange can be null)
                if (maxDisplayRange != null)
                {
                    // Convert the range boundaries to cell names (e.g., "A1")
                    string startCell = CellsHelper.CellIndexToName(maxDisplayRange.FirstRow, maxDisplayRange.FirstColumn);
                    string endCell = CellsHelper.CellIndexToName(
                        maxDisplayRange.FirstRow + maxDisplayRange.RowCount - 1,
                        maxDisplayRange.FirstColumn + maxDisplayRange.ColumnCount - 1);

                    // Define the print area using the calculated range
                    worksheet.PageSetup.PrintArea = $"{startCell}:{endCell}";
                }

                // Save the workbook with the defined print area
                workbook.Save("PrintAreaFromMaxDisplayRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
