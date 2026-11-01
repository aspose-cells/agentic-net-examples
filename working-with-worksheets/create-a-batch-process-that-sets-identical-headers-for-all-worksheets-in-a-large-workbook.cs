// Title: How to set the same left header on every worksheet in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set left header for all worksheets in a workbook | batch apply identical page header to each sheet with Aspose.Cells .NET | loop through worksheets and set page setup header using Aspose.Cells | update worksheet headers in a large Excel file programmatically Aspose.Cells | save workbook after modifying headers across multiple sheets in C#
// Tags: set left header Aspose.Cells page setup | iterate workbook worksheets C# Aspose.Cells | batch header update Excel workbook | dynamic pageSetup property handling Aspose.Cells | save modified workbook after header change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, iterates over every worksheet, and assigns the same left header text ("Confidential Report") to each sheet using a dynamic PageSetup reference. After updating all sheets, the workbook is saved to a new file, with error handling for missing files and unavailable header properties.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string sourcePath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the source file exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Header text to apply to every worksheet
            string headerText = "Confidential Report";

            // Apply the header to each worksheet using dynamic to avoid compile‑time binding issues
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    dynamic pageSetup = sheet.PageSetup;
                    pageSetup.LeftHeader = headerText; // Set left header
                }
                catch (Exception ex)
                {
                    // Log if the header property is unavailable
                    Console.WriteLine($"Unable to set header for sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
