// Title: Clear printer settings for all worksheets in an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# clear printer settings on each worksheet | remove printer configuration from all sheets in an Excel file using Aspose.Cells | set PageSetup.PrinterSettings to null for every worksheet in .NET | how to reset printer data in an Excel workbook with Aspose.Cells | iterate through workbook worksheets to delete printer settings C#
// Tags: worksheet printer data removal Aspose.Cells | nullify PageSetup.PrinterSettings .NET | excel workbook printer settings purge | loop over worksheets to clear printer data | Aspose.Cells printer settings cleanup

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, iterates over each worksheet, assigns null to the PageSetup.PrinterSettings property to remove any printer configuration, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Clear printer settings for each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // PageSetup.PrinterSettings can be set to null to remove any printer configuration
                sheet.PageSetup.PrinterSettings = null;
            }

            // Save the modified workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
