// Title: Clone Worksheet PageSetup from a Template Sheet using Aspose.Cells for .NET (C#)
// Description: Loads a template workbook, extracts its first worksheet's PageSetup, creates a new workbook, adds a fresh sheet, copies the template's print settings with CopyOptions, and saves the result. Handles missing template by generating a simple placeholder.
// Keywords: Aspose.Cells C# | copy PageSetup | clone worksheet print settings | Aspose.Cells CopyOptions | Excel page layout automation .NET | duplicate margins orientation scaling | template worksheet page setup
// Common Searches: Aspose.Cells copy page setup between worksheets | C# clone worksheet print settings using Aspose | How to duplicate PageSetup in Aspose.Cells | Copy margins and orientation from template sheet Aspose.Cells | Aspose.Cells CopyOptions example for PageSetup
// Developer Intent: Replicate the PageSetup configuration of a template sheet and apply it to a newly created worksheet.
// Use Cases: Generate reports that share identical print margins, orientation, and scaling with a predefined template. | Programmatically add multiple sheets that need consistent printing layouts across a workbook. | Migrate existing workbook page‑setup settings to a new workbook during automated document creation.
// AI Prompts: Provide C# code that copies the full PageSetup (including headers, footers, and print options) from a template worksheet to another worksheet using Aspose.Cells. | Show an example of applying the cloned PageSetup to several new sheets in a loop with Aspose.Cells for .NET. | Explain how CopyOptions influences the PageSetup copy operation in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads a template workbook, extracts its first worksheet's PageSetup, creates a new workbook, adds a fresh sheet, copies the template's print settings with CopyOptions, and saves the result. Handles missing template by generating a simple placeholder.
class ClonePageSetupExample
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the template file exists; create a simple one if it does not.
            if (!File.Exists(templatePath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "TemplateSheet";
                tempWb.Save(templatePath);
            }

            // Load the template workbook containing the source worksheet.
            Workbook templateWorkbook = new Workbook(templatePath);
            Worksheet templateSheet = templateWorkbook.Worksheets[0]; // assume first sheet is the template

            // Create a new workbook for the cloned sheet.
            Workbook newWorkbook = new Workbook();
            // Remove the default sheet created with the new workbook.
            newWorkbook.Worksheets.Clear();

            // Add a new worksheet to the new workbook.
            Worksheet newSheet = newWorkbook.Worksheets.Add("ClonedSheet");

            // Clone the page setup from the template worksheet to the new worksheet.
            newSheet.PageSetup.Copy(templateSheet.PageSetup, new CopyOptions());

            // Save the workbook with the cloned page setup.
            newWorkbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
