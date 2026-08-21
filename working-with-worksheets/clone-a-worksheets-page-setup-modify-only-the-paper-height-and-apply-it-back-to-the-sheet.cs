// Title: C# – Clone Worksheet PageSetup, Adjust Paper Height, and Apply to Another Sheet with Aspose.Cells
// Description: Demonstrates how to copy the full PageSetup from a source worksheet to a target worksheet using Aspose.Cells for .NET, then modify only the paper height while preserving the original width via CustomPaperSize, and finally save the workbook.
// Keywords: Aspose.Cells | C# | .NET | PageSetup.Copy | Clone worksheet page setup | CustomPaperSize | modify paper height | worksheet printing settings | CopyOptions | page layout automation
// Common Searches: Aspose.Cells copy page setup between worksheets | change only paper height after cloning page setup C# | custom paper size preserving width Aspose.Cells | how to use PageSetup.Copy and CustomPaperSize | duplicate worksheet layout and adjust paper dimensions
// Developer Intent: Copy a worksheet’s PageSetup, change just the paper height, and apply the updated layout to another sheet.
// Use Cases: Generate multi‑sheet reports where all sheets share margins, orientation, and scaling, but each sheet needs a different paper height for varied content length. | Automate the duplication of page‑layout settings across worksheets while customizing printable area without manual re‑configuration. | Create printable forms that inherit a standard layout from a template sheet and then adjust height to accommodate extra rows or graphics.
// AI Prompts: Write C# code with Aspose.Cells that copies a worksheet’s PageSetup to another sheet and sets a custom paper size keeping the original width while increasing the height. | Explain the interaction between PageSetup.Copy and CustomPaperSize for modifying only the paper height in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for cloning page layout settings and adjusting paper dimensions without affecting other PageSetup properties.

using System;
using Aspose.Cells;

// Demonstrates how to copy the full PageSetup from a source worksheet to a target worksheet using Aspose.Cells for .NET, then modify only the paper height while preserving the original width via CustomPaperSize, and finally save the workbook.
class ClonePageSetupModifyPaperHeight
{
    static void Main()
    {
        // Create a new workbook and add a second worksheet
        Workbook workbook = new Workbook();
        workbook.Worksheets.Add();

        // Source worksheet (the one we will clone from)
        Worksheet sourceSheet = workbook.Worksheets[0];
        // Target worksheet (the one we will apply the cloned settings to)
        Worksheet targetSheet = workbook.Worksheets[1];

        // Example: set an initial page setup on the source sheet
        sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

        // Clone the entire page setup from source to target
        targetSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

        // Modify only the paper height on the target sheet
        // Retrieve the current paper width (read‑only)
        double currentWidth = targetSheet.PageSetup.PaperWidth;
        // Define a new height (e.g., increase by 1 inch)
        double newHeight = targetSheet.PageSetup.PaperHeight + 1.0;
        // Apply a custom paper size keeping the original width and using the new height
        targetSheet.PageSetup.CustomPaperSize(currentWidth, newHeight);

        // Save the workbook with the modified page setup
        workbook.Save("ClonedPageSetup.xlsx");
    }
}
