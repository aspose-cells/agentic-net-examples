// Title: Set a uniform StandardHeight for every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to loop through all worksheets in a .NET workbook, assign the same Cells.StandardHeight value, and save the updated file. | Programmatically enforce a consistent default row height across an entire Excel file by setting Worksheet.Cells.StandardHeight for each sheet with C#. | Apply a global row height of 15 points to every sheet in an existing workbook using Aspose.Cells and output a new workbook.
// Common Searches: Aspose.Cells C# set same row height for all worksheets in a workbook | How to change default row height for every sheet using Aspose.Cells .NET | Batch update Excel row height across multiple worksheets with Aspose.Cells | Set Cells.StandardHeight property for each worksheet in C# | Apply uniform row height to an existing Excel file using Aspose.Cells
// Tags: Worksheet.Cells.StandardHeight property Aspose.Cells | global row height for all sheets .NET | batch set default row height Excel workbook C# | uniform row height across worksheets using Aspose.Cells | standard row height 15 points Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The C# program loads 'input.xlsx', checks its existence, iterates through each worksheet, sets the Cells.StandardHeight to 15 points, saves the modified workbook as 'output.xlsx', and handles any runtime exceptions.
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
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Desired standard row height (in points)
            double standardHeight = 15.0;

            // Apply the standard height to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set the default row height for the sheet
                sheet.Cells.StandardHeight = standardHeight;
            }

            // Save the modified workbook
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
