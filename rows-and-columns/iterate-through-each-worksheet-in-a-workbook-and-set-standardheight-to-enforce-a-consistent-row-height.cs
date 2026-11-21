// Title: How to set a uniform standard row height for all worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set standard row height for every sheet in a workbook | C# iterate through worksheets and apply same row height with Aspose.Cells | Set default row height across all worksheets using Aspose.Cells .NET | How to change Cells.StandardHeight for multiple worksheets in Aspose.Cells | Batch update row height in Excel file with Aspose.Cells for .NET
// Tags: Aspose.Cells set worksheet standard row height | C# batch row height update Excel | Iterate worksheets Cells.StandardHeight | Uniform row height Aspose.Cells .NET | Excel workbook default row height C#

using Aspose.Cells;
using System;
using System.IO;

// Loads an existing Excel file, defines a 15‑point standard row height, loops through each worksheet in the workbook, assigns the height via the Cells.StandardHeight property, and saves the modified workbook, including file existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            var workbook = new Workbook(inputPath);

            // Desired standard row height (in points)
            double standardHeight = 15.0;

            // Iterate through each worksheet and set the default row height
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Correct API: set the StandardHeight via the Cells collection
                sheet.Cells.StandardHeight = standardHeight;
            }

            // Save the updated workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
