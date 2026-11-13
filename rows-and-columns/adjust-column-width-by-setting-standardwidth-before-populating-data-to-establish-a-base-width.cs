// Title: Set a default column width for all columns with Cells.StandardWidth in Aspose.Cells for .NET before writing data
// Common Searches: Aspose.Cells set default column width for entire worksheet in C# | How to apply a base column width using Cells.StandardWidth before populating cells | C# Aspose.Cells initialize column widths prior to adding data | Set uniform column width in a new Excel workbook with Aspose.Cells .NET | StandardWidth property usage example for column sizing in Aspose.Cells
// Tags: Cells.StandardWidth column sizing | default worksheet column width Aspose.Cells | initialize column width before data insertion C# | uniform column width Excel generation Aspose.Cells | set base column width .NET Excel library

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, set a base column width of 20 characters for all columns using the Cells.StandardWidth property before any data is written, add sample headers and rows, ensure the output directory exists, and save the file as Output.xlsx while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set the standard column width (in characters) before adding data
            // Use Cells.StandardWidth as Worksheet does not expose this property directly
            sheet.Cells.StandardWidth = 20; // Adjust this value to set the base width for all columns

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
