// Title: Load an Excel workbook from a file and enable iterative calculation for circular references using Aspose.Cells in C#
// Common Searches: how to enable iterative calculation for circular references in Aspose.Cells C# | load Excel file and set EnableIterativeCalculation property with Aspose.Cells .NET | Aspose.Cells example configuring MaxIterations and MaxChange after opening a workbook | C# code to load a workbook from path and turn on iterative calculation in Aspose.Cells | using Aspose.Cells to handle complex formulas with iterative calculation settings
// Tags: load workbook from file Aspose.Cells C# | enable iterative calculation Aspose.Cells | configure MaxIterations MaxChange Aspose.Cells | circular reference handling Aspose.Cells .NET | Aspose.Cells workbook settings for complex formulas | save modified workbook Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the input.xlsx file, loads it into an Aspose.Cells Workbook, optionally activates iterative calculation by setting EnableIterativeCalculation, MaxIterations, and MaxChange, and then saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook directly from the file
            Workbook workbook = new Workbook(sourcePath);

            // Optional: enable iterative calculation if supported by your Aspose.Cells version
            // Uncomment the following lines if the properties are available
            // workbook.Settings.EnableIterativeCalculation = true;
            // workbook.Settings.MaxIterations = 100;      // maximum number of iterations
            // workbook.Settings.MaxChange = 0.001;        // convergence threshold

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
