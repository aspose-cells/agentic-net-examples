// Title: Copy a worksheet to a specific index with WorksheetCollection.AddCopy in Aspose.Cells for .NET (C#)
// Description: Shows how to duplicate a worksheet by its index using Workbook.Worksheets.AddCopy, rename the new sheet, reposition it with Worksheet.MoveTo, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | WorksheetCollection.AddCopy | copy worksheet by index | move worksheet position | duplicate sheet | insert copied sheet | Worksheet.MoveTo
// Common Searches: Aspose.Cells copy worksheet to specific position | WorksheetCollection AddCopy overload example | How to move copied worksheet in Aspose.Cells | Insert duplicated sheet at index C# | Copy sheet and set order Aspose.Cells
// Developer Intent: The developer wants to duplicate an existing worksheet and place the copy at a chosen index within the workbook using Aspose.Cells for .NET.
// Use Cases: Generate multiple template sheets and insert each copy sequentially between other worksheets. | Copy a data sheet, rename it, and position it immediately after a summary sheet for automated reporting. | Duplicate a configuration worksheet and insert it right after the main sheet to keep versioned settings together.
// AI Prompts: Write C# code that copies the worksheet at index 2 and inserts the copy at index 0 using Aspose.Cells. | Explain how Worksheet.MoveTo can reorder worksheets after using WorksheetCollection.AddCopy in a .NET workbook. | Provide a step‑by‑step guide to duplicate a sheet and place it between two existing sheets with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to duplicate a worksheet by its index using Workbook.Worksheets.AddCopy, rename the new sheet, reposition it with Worksheet.MoveTo, and save the workbook.
    public class DuplicateSheetAtPosition
    {
        public static void Main()
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Prepare the source worksheet (index 0)
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";
            sourceSheet.Cells["A1"].PutValue("Data in source sheet");

            // Duplicate the source worksheet using AddCopy overload with source index
            int copiedIndex = workbook.Worksheets.AddCopy(0); // copies sheet at index 0
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "CopiedSheet";

            // Move the copied sheet to the desired position (e.g., index 1)
            copiedSheet.MoveTo(1);

            // Save the workbook
            workbook.Save("DuplicatedSheetAtPosition.xlsx");
        }
    }
}
