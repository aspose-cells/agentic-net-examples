// Title: Hide the first worksheet tab, display formulas on the second sheet, and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing Excel file, set the first worksheet's IsVisible property to false, enable ShowFormulas on the second worksheet, and save the updated workbook to a new file using Aspose.Cells in C#. | Using Aspose.Cells for .NET, programmatically hide the tab of the first sheet, turn on formula view for the second sheet, and write the modified workbook to a specified output path.
// Common Searches: Aspose.Cells C# hide first worksheet tab programmatically | show formulas instead of values on a specific sheet with Aspose.Cells .NET | save workbook after changing sheet visibility using Aspose.Cells in C# | how to set ShowFormulas property for a worksheet in Aspose.Cells | C# example to hide a sheet tab and display formulas on another sheet with Aspose.Cells
// Tags: worksheet.IsVisible false Aspose.Cells | worksheet.ShowFormulas true C# | modify sheet tab visibility Aspose.Cells .NET | display formulas on worksheet Aspose.Cells | write updated workbook to file C#

using Aspose.Cells;
using System;
using System.IO;

// // This C# program loads an existing Excel workbook, hides the first worksheet tab, enables formula display on the second worksheet (if present), and saves the modified file to a new location using Aspose.Cells for .NET.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Hide the first worksheet tab
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.IsVisible = false; // Tab will be hidden

            // Show formulas on the second worksheet (if it exists)
            if (workbook.Worksheets.Count > 1)
            {
                Worksheet secondSheet = workbook.Worksheets[1];
                secondSheet.ShowFormulas = true; // Display formulas instead of values
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
