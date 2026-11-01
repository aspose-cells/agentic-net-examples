// Title: Add a ProfitMargin calculated field (Profit ÷ Revenue) to an existing PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: how to create a profit margin calculated field in a pivot table with Aspose.Cells C# | Aspose.Cells add calculated field Profit/Revenue to an existing pivot table programmatically | C# code to insert a data field that calculates profit margin in an Excel pivot table using Aspose | using Aspose.Cells .NET to modify a pivot table and add a custom profit margin calculation | example of adding a calculated field to a pivot table in an Aspose.Cells workbook
// Tags: calculated field addition using Aspose.Cells API | profit margin computation in Excel pivot via Aspose | programmatic pivot table modification .NET | Aspose.Cells pivot table data field API | C# Excel pivot calculated field example

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads an existing workbook, checks for a pivot table on the first worksheet, adds a calculated field named "ProfitMargin" with the formula "Profit/Revenue" directly into the data area, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Define the calculated field formula: Profit divided by Revenue
            string formula = "Profit/Revenue";

            // Add the calculated field named "ProfitMargin" and place it directly in the data area
            // The overload with the third parameter adds the field as a data field automatically
            pivotTable.AddCalculatedField("ProfitMargin", formula, true);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
