// Title: How to set a worksheet's top margin to 5 points with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet top margin to 5 points | programmatically change Excel top margin using Aspose.Cells .NET | PageSetup.TopMargin property example for 5 point margin | adjust Excel page layout margins in C# with Aspose.Cells | set top margin points for first worksheet in Aspose.Cells workbook
// Tags: worksheet top margin points Aspose.Cells | page setup top margin .NET | adjust Excel top margin programmatically | set worksheet page margins C# | modify worksheet top margin using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file with Aspose.Cells, sets the first worksheet's top margin to 5 points via the PageSetup.TopMargin property, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the top margin (in points) for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.TopMargin = 5; // points

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
