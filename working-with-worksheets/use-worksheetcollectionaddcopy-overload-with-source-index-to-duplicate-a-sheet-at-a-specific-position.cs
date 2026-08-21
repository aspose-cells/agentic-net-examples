// Title: Copy a worksheet and set its tab order with WorksheetCollection.AddCopy in Aspose.Cells for .NET
// Description: Shows how to create a workbook, rename the default sheet, duplicate it using Worksheets.AddCopy(sourceIndex), obtain the new sheet’s index, rename the copy, and move it to a specific tab position with Worksheet.MoveTo before saving.
// Keywords: Aspose.Cells | WorksheetCollection.AddCopy | duplicate worksheet C# | move worksheet | set tab order | reorder worksheets programmatically | copy sheet by index | Aspose.Cells .NET example
// Common Searches: Aspose.Cells copy worksheet to specific position | WorksheetCollection AddCopy usage example | How to move a copied sheet in Aspose.Cells | C# duplicate sheet and set tab order | Reorder worksheets programmatically Aspose.Cells
// Developer Intent: Copy an existing worksheet and insert the copy at a chosen tab index within the workbook.
// Use Cases: Create a template sheet, duplicate it, and place the copy as the second tab for a report. | Generate monthly report tabs by copying a base sheet and inserting each copy after a summary sheet. | Programmatically reorder sheets after copying to match a predefined layout before export.
// AI Prompts: Provide C# code that uses Aspose.Cells to copy a worksheet by its index and insert the copy at position 2. | Explain how Worksheet.MoveTo works after using Worksheets.AddCopy in Aspose.Cells. | Show an example of duplicating multiple worksheets and arranging them in a custom order with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, rename the default sheet, duplicate it using Worksheets.AddCopy(sourceIndex), obtain the new sheet’s index, rename the copy, and move it to a specific tab position with Worksheet.MoveTo before saving.
    public class DuplicateSheetAtSpecificPosition
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
            // Create a new workbook (default workbook contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the original worksheet (index 0) and add some data
            Worksheet originalSheet = workbook.Worksheets[0];
            originalSheet.Name = "Original";
            originalSheet.Cells["A1"].PutValue("Data in the original sheet");

            // Duplicate the original worksheet using AddCopy overload that takes the source index
            // This method returns the index of the newly created copy (added at the end)
            int copiedIndex = workbook.Worksheets.AddCopy(0);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "Copied";

            // Move the copied sheet to the desired position, e.g., index 1 (second tab)
            // MoveTo repositions the sheet within the workbook's sheet collection
            copiedSheet.MoveTo(1);

            // Save the workbook to a file
            string outputPath = "DuplicatedSheetAtPosition.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
