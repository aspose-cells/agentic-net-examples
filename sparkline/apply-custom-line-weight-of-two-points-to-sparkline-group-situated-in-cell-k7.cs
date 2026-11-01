// Title: How to set a 2‑point custom line weight for a sparkline group that includes cell K7 using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that finds the sparkline group covering cell K7 and sets ShowCustomLineWeight to true and CustomLineWeight to 2. | Show an example of iterating over Worksheet.SparklineGroups in Aspose.Cells to apply a two‑point custom line weight to the group whose LocationRange contains K7.
// Common Searches: aspnet cells c# set sparkline custom line weight for cell K7 | how to change sparkline line thickness in an Excel file using Aspose.Cells | apply two point line weight to sparkline group located at K7 with Aspose.Cells .NET | C# code to enable custom line weight for sparkline groups in an Aspose.Cells workbook
// Tags: Aspose.Cells set sparkline custom line weight | C# sparkline group line thickness | modify sparkline line weight in Excel workbook | Aspose.Cells SparklineGroups iteration | custom line weight for sparkline in .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, locates any sparkline group that includes cell K7, enables a custom line weight, sets it to 2 points, and saves the updated file.
class SparklineWeightExample
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
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Zero‑based indices for cell K7
            int targetRow = 6;      // Row 7 in Excel
            int targetColumn = 10;  // Column K in Excel

            // Attempt to process sparkline groups using dynamic to avoid compile‑time dependency
            try
            {
                dynamic sheetDyn = sheet;
                foreach (var group in sheetDyn.SparklineGroups)
                {
                    // Determine if the group's location range includes cell K7
                    var loc = group.LocationRange;
                    bool containsK7 = targetRow >= loc.StartRow && targetRow <= loc.EndRow &&
                                      targetColumn >= loc.StartColumn && targetColumn <= loc.EndColumn;

                    if (containsK7)
                    {
                        // Enable custom line weight and set it to 2 points
                        group.ShowCustomLineWeight = true;
                        group.CustomLineWeight = 2;
                    }
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // Sparkline feature not available in this version of Aspose.Cells
                Console.WriteLine("Sparkline functionality is not supported by the current Aspose.Cells library.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
