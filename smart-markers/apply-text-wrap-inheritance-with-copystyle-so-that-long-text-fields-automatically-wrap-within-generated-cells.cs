using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsWrapInheritanceDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // -----------------------------------------------------------------
                // 1. Define a source style with text wrapping enabled
                // -----------------------------------------------------------------
                Style wrapStyle = workbook.CreateStyle();
                wrapStyle.IsTextWrapped = true; // enable wrap
                // Apply the style to a single source cell (A1)
                cells["A1"].SetStyle(wrapStyle);

                // -----------------------------------------------------------------
                // 2. Create source and destination ranges
                // -----------------------------------------------------------------
                // Source range containing the wrap style (A1)
                AsposeRange srcRange = cells.CreateRange("A1");
                // Destination range where long text will be placed (B2:D4)
                AsposeRange destRange = cells.CreateRange("B2:D4");

                // -----------------------------------------------------------------
                // 3. Copy the wrap style from the source range to the destination range
                // -----------------------------------------------------------------
                destRange.CopyStyle(srcRange);

                // -----------------------------------------------------------------
                // 4. Populate the destination range with long text
                // -----------------------------------------------------------------
                string longText = "This is a very long text that should be wrapped automatically in each cell of the range. " +
                                  "It contains multiple sentences to ensure wrapping works correctly and demonstrates inheritance of the wrap style.";
                // Use overload that specifies conversion and formula flags (both false)
                destRange.PutValue(longText, false, false);

                // -----------------------------------------------------------------
                // 5. Adjust column widths so that wrapping becomes visible
                // -----------------------------------------------------------------
                // Set column widths for columns B, C and D (indexes 1,2,3)
                cells.SetColumnWidth(1, 20);
                cells.SetColumnWidth(2, 20);
                cells.SetColumnWidth(3, 20);

                // -----------------------------------------------------------------
                // 6. Auto‑fit rows to display the wrapped text properly
                // -----------------------------------------------------------------
                worksheet.AutoFitRows();

                // -----------------------------------------------------------------
                // 7. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "WrapInheritanceOutput.xlsx";

                // Ensure we can write to the target location
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}