// Title: Create an Average Price calculated field in an existing PivotTable with Aspose.Cells for .NET (C#)
// Common Searches: how to add a calculated field that divides total sales by quantity in an Aspose.Cells PivotTable using C# | Aspose.Cells C# example for adding average price formula to a PivotTable | C# code to insert a custom calculated field into an existing PivotTable with Aspose.Cells | using Aspose.Cells to compute average price in a PivotTable from Total Sales and Quantity columns | add calculated field to PivotTable in Aspose.Cells workbook programmatically
// Tags: add calculated field to PivotTable Aspose.Cells | average price formula in PivotTable C# | custom calculated field Aspose.Cells .NET | modify existing PivotTable programmatically | Aspose.Cells workbook pivot table calculation

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// Loads a workbook, checks for a PivotTable on the first worksheet, defines the formula ='Total Sales'/'Quantity' as an Average Price calculated field, adds it to the PivotTable, and saves the updated file.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTables found in the first worksheet.");
                return;
            }

            // Access the first PivotTable in the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Define the formula for the calculated field: Average Price = Total Sales / Quantity
            // Field names are enclosed in single quotes if they contain spaces.
            string formula = "='Total Sales'/'Quantity'";

            // Add the calculated field to the PivotTable
            pivotTable.AddCalculatedField("Average Price", formula);

            // Save the workbook with the new calculated field
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
