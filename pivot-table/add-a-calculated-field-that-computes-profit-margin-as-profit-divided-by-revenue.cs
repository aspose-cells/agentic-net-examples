// Title: Add a ProfitMargin calculated field (Profit/Revenue) to an existing PivotTable using Aspose.Cells for .NET C#
// Common Searches: C# Aspose.Cells add calculated field ProfitMargin to existing pivot table | How to create a profit margin field in Excel pivot using Aspose.Cells library | Aspose.Cells .NET example for adding custom calculation to pivot table | Add Profit/Revenue calculated field to pivot table in C# workbook | Using Aspose.Cells to modify pivot tables and include profit margin formula
// Tags: aspnet cells pivot table calculated field addition | margin calculation field integration aspnet cells | c# update pivot table using aspnet cells | excel workbook pivot table profit revenue field | aspnet cells add profit revenue field to pivot

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The example loads an existing Excel workbook, checks for a pivot table on the first worksheet, adds a calculated field named "ProfitMargin" with the formula "Profit/Revenue" to that pivot table, and saves the updated workbook to a new file.
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

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables == null || sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Retrieve the first pivot table on the worksheet
            PivotTable pivotTable = sheet.PivotTables[0];

            // Add a calculated field named "ProfitMargin" with the formula Profit/Revenue
            // Use AddCalculatedField method (available in Aspose.Cells)
            pivotTable.AddCalculatedField("ProfitMargin", "Profit/Revenue");

            // Save the workbook with the new calculated field
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
