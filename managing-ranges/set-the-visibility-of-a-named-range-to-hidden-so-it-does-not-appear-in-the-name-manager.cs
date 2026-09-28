// Title: Hide a specific named range from the Name Manager in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to retrieve a named range by its name, set its IsVisible property to false, and save the workbook. | Write C# code that loads an existing .xlsx file, hides a given named range from the Name Manager, and writes the result to a new file with Aspose.Cells. | Programmatically make a named range invisible in Excel using Aspose.Cells for .NET, handling missing files and ensuring the output directory exists.
// Common Searches: how to hide a named range from name manager using Aspose.Cells C# | Aspose.Cells set named range visibility false example | C# hide Excel named range programmatically with Aspose.Cells | remove named range from name manager without deleting it Aspose.Cells | save workbook after changing named range visibility Aspose.Cells .NET
// Tags: Aspose.Cells hide named range | C# set named range visibility | Excel Name Manager invisible entry | Aspose.Cells modify name visibility | save workbook after hiding named range

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, locates the named range 'MyNamedRange', sets its IsVisible property to false so it no longer appears in the Name Manager, and saves the modified file to a new location, including checks for file existence and error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range by its name
            Name namedRange = workbook.Worksheets.Names["MyNamedRange"];
            if (namedRange != null)
            {
                // Hide the named range from the Name Manager
                namedRange.IsVisible = false;
            }
            else
            {
                Console.WriteLine("Named range 'MyNamedRange' not found.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the updated visibility setting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
