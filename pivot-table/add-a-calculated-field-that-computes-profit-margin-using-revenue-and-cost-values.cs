// Title: Add a Profit Margin calculated field to an existing PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: how to add a calculated field named ProfitMargin to a PivotTable with Aspose.Cells C# | Aspose.Cells C# add custom formula (Revenue-Cost)/Revenue to pivot table | programmatically insert profit margin calculation into Excel pivot using Aspose.Cells | C# code to create a calculated field in an existing workbook pivot table with Aspose | Aspose.Cells add calculated field to a pivot table without rebuilding the pivot
// Tags: Aspose.Cells calculated field pivot table | profit margin formula pivot C# | custom calculation Excel pivot Aspose.Cells | modify existing pivot table Aspose.Cells | C# add field to Excel pivot using Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing workbook, checks for a pivot table on the first worksheet, adds a calculated field called "ProfitMargin" with the formula "(Revenue-Cost)/Revenue", and saves the updated file.
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure that at least one pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Retrieve the first pivot table (adjust index if there are multiple)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a calculated field named "ProfitMargin"
            // Formula: (Revenue - Cost) / Revenue
            // Column names must match the field names used in the pivot table source data
            pivotTable.AddCalculatedField("ProfitMargin", "(Revenue-Cost)/Revenue");

            // Save the workbook with the new calculated field
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
