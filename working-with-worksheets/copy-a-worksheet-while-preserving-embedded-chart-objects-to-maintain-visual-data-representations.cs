// Title: How to copy an Excel worksheet and keep its embedded charts using Aspose.Cells for .NET
// Common Searches: C# copy worksheet with charts using Aspose.Cells AddCopy method | duplicate Excel sheet while preserving embedded chart objects in Aspose.Cells | Aspose.Cells copy sheet retain chart objects example | copy worksheet and keep visual data representations Aspose.Cells .NET | how to clone a worksheet with charts in Aspose.Cells C#
// Tags: AddCopy worksheet method Aspose.Cells | copy worksheet with embedded charts C# | duplicate Excel sheet preserving chart objects | Aspose.Cells worksheet cloning charts | preserve visual data when copying sheet Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source workbook, verifies the file, selects a worksheet, and uses the AddCopy overload to duplicate the sheet while retaining all embedded chart objects. It then renames the copied sheet, ensures the output directory exists, and saves the workbook, handling any runtime errors.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The source file \"{sourcePath}\" was not found.");
            return;
        }

        try
        {
            // Load the source workbook that contains the worksheet with embedded charts
            Workbook workbook = new Workbook(sourcePath);

            // Reference the worksheet you want to copy (by name or index)
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name if different
            if (sourceSheet == null)
            {
                Console.WriteLine("Error: The specified worksheet \"Sheet1\" does not exist in the source workbook.");
                return;
            }

            // Copy the worksheet, including all embedded objects such as charts
            // Use the overload that accepts the worksheet name (compatible with all Aspose.Cells versions)
            int copiedSheetIndex = workbook.Worksheets.AddCopy(sourceSheet.Name);

            // Optionally rename the copied worksheet to avoid name conflicts
            Worksheet copiedSheet = workbook.Worksheets[copiedSheetIndex];
            copiedSheet.Name = sourceSheet.Name + "_Copy";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook; the copied sheet retains its charts and visual data representations
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
