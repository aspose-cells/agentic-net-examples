// Title: Copy an Excel range to a new workbook and freeze the top row using Aspose.Cells for .NET (C#)
// Description: C# example that loads a source workbook, copies a defined range (e.g., A1:C5) into a freshly created workbook, applies FreezePanes to lock the first row, and saves the result as a separate file.
// Keywords: Aspose.Cells copy range C# | freeze top row Aspose.Cells | Aspose.Cells FreezePanes example | copy Excel range between workbooks .NET | Aspose.Cells range copy and freeze | C# Excel automation Aspose | create new workbook from range
// Common Searches: copy range from one Excel file to another using Aspose.Cells .NET | freeze first row after copying data with Aspose.Cells | Aspose.Cells programmatically copy cells and freeze panes | C# example to copy a block of cells to a new workbook | how to use FreezePanes in Aspose.Cells
// Developer Intent: Copy a specific cell block to a new workbook and keep the header row fixed while scrolling.
// Use Cases: Generate a lightweight report that contains only the required data slice while keeping column headers visible. | Distribute a portion of a large worksheet to external partners with the top row frozen for readability. | Automate extraction of a template section into a separate file for downstream processing, preserving the header row.
// AI Prompts: Provide C# code with Aspose.Cells to copy range A1:D10 from source.xlsx to a new workbook and freeze the first two rows. | Show an Aspose.Cells .NET snippet that copies a dynamic range based on used cells to a new file and applies FreezePanes to lock the header row. | Explain how to copy multiple non‑contiguous ranges into a new workbook and set FreezePanes for the top row using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyAndFreeze
{
    // C# example that loads a source workbook, copies a defined range (e.g., A1:C5) into a freshly created workbook, applies FreezePanes to lock the first row, and saves the result as a separate file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsx";
                const string destPath = "copied_and_frozen.xlsx";

                // Ensure the source file exists; if not, create a simple workbook for demo purposes
                if (!File.Exists(sourcePath))
                {
                    var tempWb = new Workbook();
                    var tempSheet = tempWb.Worksheets[0];
                    tempSheet.Cells["A1"].PutValue("Demo");
                    tempSheet.Cells["B2"].PutValue(123);
                    tempSheet.Cells["C3"].PutValue(DateTime.Now);
                    tempWb.Save(sourcePath);
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Define the source range to copy (e.g., A1:C5)
                AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C5");

                // Create a new workbook that will receive the copied range
                Workbook destWorkbook = new Workbook(); // creates a default worksheet
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define the destination range in the new workbook (starting at A1)
                AsposeRange destRange = destSheet.Cells.CreateRange("A1:C5");

                // Copy the source range into the destination range
                sourceRange.Copy(destRange);

                // Freeze the top row in the destination worksheet view
                // Freeze at row index 1 (second row), column index 0, freezing 1 row and 0 columns
                destSheet.FreezePanes(1, 0, 1, 0);

                // Save the resulting workbook
                destWorkbook.Save(destPath);
                Console.WriteLine($"Workbook saved successfully to '{destPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
