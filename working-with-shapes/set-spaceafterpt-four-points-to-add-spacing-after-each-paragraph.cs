// Title: Set a uniform row height of 20 points for all worksheets with Aspose.Cells for .NET (C#)
// Common Searches: how to set the same row height for every sheet using Aspose.Cells C# | Aspose.Cells C# set StandardHeight property for all worksheets | apply uniform row height across an Excel workbook with Aspose.Cells | C# code to change row height to 20 points in each worksheet using Aspose.Cells
// Tags: standard row height property Aspose.Cells | set uniform worksheet row height C# | apply row height to all sheets Aspose.Cells | modify Excel row height programmatically .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, iterates through each worksheet, sets each sheet's StandardHeight to 20 points to achieve a consistent row height, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Example operation: set a uniform row height (e.g., 20 points)
                sheet.Cells.StandardHeight = 20;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
