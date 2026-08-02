// Title: Copy worksheet PageSetup to multiple sheets with Aspose.Cells for .NET
// Description: Loads an Excel file, extracts the PageSetup from a designated source sheet, then iterates through the remaining worksheets, applying the same print layout using PageSetup.Copy and saving the result.
// Keywords: Aspose.Cells PageSetup copy | C# copy print settings between worksheets | duplicate page layout Excel | Aspose.Cells copy options | standardize worksheet margins .NET | transfer header footer settings | Excel workbook page configuration | Aspose.Cells loop example | copy worksheet print area
// Common Searches: Aspose.Cells copy page setup from one sheet to others | C# loop to apply same print settings to all worksheets | How to duplicate worksheet margins using Aspose | Copy orientation and scaling across Excel sheets .NET | Aspose.Cells PageSetup.Copy usage example
// Developer Intent: Apply the PageSetup of a single worksheet to every other sheet in the same workbook.
// Use Cases: Ensure consistent print margins, orientation, and scaling before exporting a multi‑sheet workbook to PDF. | Synchronize header/footer content across all worksheets for uniform printed reports. | Batch‑apply page layout settings when generating Excel files for regional compliance or corporate branding.
// AI Prompts: Write C# code that copies the PageSetup from a chosen worksheet to a list of target worksheets, including null checks and customizable CopyOptions. | Show how to copy PageSetup and also set an identical print area for each destination sheet using Aspose.Cells. | Explain how to duplicate page layout without altering data validation, conditional formatting, or existing cell values in the target worksheets.

using System;
using Aspose.Cells;

// Loads an Excel file, extracts the PageSetup from a designated source sheet, then iterates through the remaining worksheets, applying the same print layout using PageSetup.Copy and saving the result.
class Program
{
    static void Main()
    {
        // Load the workbook containing the source and target worksheets
        Workbook workbook = new Workbook("input.xlsx");

        // Define the source worksheet (e.g., the first worksheet)
        int sourceIndex = 0;
        PageSetup sourcePageSetup = workbook.Worksheets[sourceIndex].PageSetup;

        // Loop through all worksheets and copy the page setup to each target worksheet
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            // Skip the source worksheet itself
            if (i == sourceIndex) continue;

            // Copy the page setup settings from the source worksheet to the current worksheet
            workbook.Worksheets[i].PageSetup.Copy(sourcePageSetup, new CopyOptions());
        }

        // Save the workbook with the updated page setup settings
        workbook.Save("output.xlsx");
    }
}
