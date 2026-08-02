// Title: Clone worksheet PageSetup from a template to a new sheet with Aspose.Cells for .NET
// Description: Demonstrates how to load a template workbook, retrieve its first worksheet, create a new workbook, add a fresh sheet, and copy the template's PageSetup (margins, orientation, scaling, etc.) to the new sheet using Aspose.Cells' PageSetup.Copy method with default CopyOptions. Includes file‑existence checks and error handling, then saves the result.
// Keywords: Aspose.Cells PageSetup copy | clone worksheet print settings C# | copy page layout between workbooks | Aspose.Cells .NET page setup example | duplicate worksheet margins orientation
// Common Searches: Aspose.Cells copy page setup from one sheet to another | C# clone worksheet print settings using Aspose.Cells | How to duplicate PageSetup in Aspose.Cells .NET | Copy margins and orientation between Excel sheets programmatically
// Developer Intent: Copy the PageSetup of a template worksheet and apply it to a newly created worksheet in a separate workbook.
// Use Cases: Generate reports that share the same print layout as a master template. | Create multiple sheets with identical page margins, orientation, and scaling for batch printing or PDF export. | Programmatically ensure consistent printing settings when adding new worksheets to dynamically generated workbooks.
// AI Prompts: Write C# code using Aspose.Cells to copy the PageSetup from a source worksheet to a destination worksheet, including checks for missing template files and exception handling. | Explain how the CopyOptions parameter influences the PageSetup.Copy method and show examples of copying only specific properties such as margins or orientation.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load a template workbook, retrieve its first worksheet, create a new workbook, add a fresh sheet, and copy the template's PageSetup (margins, orientation, scaling, etc.) to the new sheet using Aspose.Cells' PageSetup.Copy method with default CopyOptions. Includes file‑existence checks and error handling, then saves the result.
class ClonePageSetupDemo
{
    static void Main()
    {
        try
        {
            const string templatePath = "template.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook that contains the source worksheet
            Workbook templateWorkbook = new Workbook(templatePath);
            Worksheet templateWorksheet = templateWorkbook.Worksheets[0]; // source sheet

            // Create a new workbook for the destination sheet
            Workbook destinationWorkbook = new Workbook();
            // Remove the default sheet created by the constructor
            destinationWorkbook.Worksheets.Clear();

            // Add a new worksheet where the page setup will be cloned
            int newSheetIndex = destinationWorkbook.Worksheets.Add();
            Worksheet newWorksheet = destinationWorkbook.Worksheets[newSheetIndex];
            newWorksheet.Name = "ClonedSheet";

            // Clone the page setup from the template worksheet to the new worksheet
            newWorksheet.PageSetup.Copy(templateWorksheet.PageSetup, new CopyOptions());

            // Save the workbook with the cloned page setup
            destinationWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
