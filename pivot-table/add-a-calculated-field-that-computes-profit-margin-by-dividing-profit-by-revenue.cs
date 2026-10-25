// Title: Add a ProfitMargin calculated field to an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: asp.net c# add profit margin calculated field to an existing pivot table with aspose.cells | sample code for creating a calculated field profit/revenue in an Excel pivot table using Aspose.Cells | how to programmatically insert a custom calculated field into a pivot table in C# | Aspose.Cells example for modifying a workbook to include a profit margin field in its pivot table | C# Aspose.Cells AddCalculatedField usage for Excel pivot table customization
// Tags: Aspose.Cells calculated field API | Excel pivot table custom field C# | profit margin formula pivot table | modify existing workbook Aspose.Cells | C# programmatic pivot table customization

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel file, accesses the first worksheet, retrieves the pivot table named "PivotTable1", adds a calculated field called "ProfitMargin" defined by the formula "Profit/Revenue" using Aspose.Cells, and saves the updated workbook, with error handling for missing files and exceptions.
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

            // Access the worksheet that contains the pivot table (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the pivot table by name
            PivotTable pivotTable = sheet.PivotTables["PivotTable1"];
            if (pivotTable == null)
            {
                Console.WriteLine("Pivot table 'PivotTable1' not found.");
                return;
            }

            // Add a calculated field named "ProfitMargin" that divides Profit by Revenue
            // Use the AddCalculatedField method provided by Aspose.Cells
            pivotTable.AddCalculatedField("ProfitMargin", "Profit/Revenue");

            // Save the workbook with the new calculated field
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
