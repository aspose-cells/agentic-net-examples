// Title: Duplicate an Excel worksheet and retain its pivot tables using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy worksheet with pivot tables example | how to clone an Excel sheet while keeping pivot tables in .NET | duplicate worksheet preserving pivot cache using Aspose.Cells | C# code to copy a worksheet and retain pivot tables | copy Excel worksheet including pivot tables with Aspose.Cells library
// Tags: duplicate worksheet preserving pivot tables Aspose.Cells | copy worksheet with pivot cache .NET | clone Excel sheet retaining pivot tables C# | Aspose.Cells worksheet copy including pivot tables | preserve pivot tables when duplicating worksheet

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, verifies the source sheet, adds a new worksheet, copies the source sheet (including data, styles, merged cells, and pivot tables) to the new sheet, and saves the result as a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Names of the source and duplicated worksheets
            string sourceSheetName = "Sheet1";
            string duplicatedSheetName = "Sheet1_Copy";

            // Get the source worksheet
            Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];
            if (sourceSheet == null)
            {
                Console.WriteLine($"Worksheet \"{sourceSheetName}\" does not exist in the workbook.");
                return;
            }

            // Add a new empty worksheet and set its name
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet newSheet = workbook.Worksheets[newSheetIndex];
            newSheet.Name = duplicatedSheetName;

            // Copy the source worksheet to the new worksheet (copies data, styles, merged cells, etc.)
            sourceSheet.Copy(newSheet);

            // Save the workbook with the duplicated worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
