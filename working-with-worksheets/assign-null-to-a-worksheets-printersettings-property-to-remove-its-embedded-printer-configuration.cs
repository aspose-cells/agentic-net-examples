// Title: How to clear a worksheet’s embedded printer settings with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set Worksheet.PageSetup.PrinterSettings to null | remove printer configuration from Excel sheet using Aspose.Cells | clear embedded printer settings in a .xlsx workbook with Aspose.Cells | reset worksheet printer setup programmatically in .NET | delete printer settings from Aspose.Cells worksheet before saving
// Tags: Worksheet.PageSetup.PrinterSettings null Aspose.Cells | clear embedded printer configuration .NET | remove worksheet printer settings C# | Aspose.Cells workbook without printer setup | reset Excel sheet printer configuration programmatically

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates creating or loading a Workbook, accessing its first Worksheet, and assigning null to the PageSetup.PrinterSettings property to remove any embedded printer configuration before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Remove the embedded printer configuration
            // The PrinterSettings property is accessed via PageSetup
            sheet.PageSetup.PrinterSettings = null;

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
