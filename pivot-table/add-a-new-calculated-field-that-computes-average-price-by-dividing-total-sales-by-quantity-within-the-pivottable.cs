// Title: Add an Average Price calculated field to an Excel PivotTable using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add a calculated field to an existing pivot table | Create an average price field in an Excel pivot table with Aspose.Cells | Use AddCalculatedField method to divide Total Sales by Quantity in .NET | How to define custom calculations in a PivotTable using Aspose.Cells library | C# example for adding a custom calculated field to a workbook pivot
// Tags: Aspose.Cells AddCalculatedField pivot table | C# average price calculated field Excel | custom pivot table calculation Aspose.Cells | divide total sales quantity calculated field | Excel pivot table custom formula .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The sample loads an existing workbook, accesses the first worksheet, retrieves the pivot table named "PivotTable1", defines a formula that divides the 'Total Sales' field by the 'Quantity' field, adds a new calculated field called "AveragePrice" using the AddCalculatedField method, and saves the updated workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the pivot table by its name (ensure the name matches the workbook)
            PivotTable pivotTable = worksheet.PivotTables["PivotTable1"];
            if (pivotTable == null)
            {
                Console.WriteLine("Pivot table 'PivotTable1' not found.");
                return;
            }

            // Define the formula for the calculated field: Total Sales divided by Quantity
            string formula = "='Total Sales'/'Quantity'";

            // Add the calculated field named "AveragePrice" to the pivot table
            pivotTable.AddCalculatedField("AveragePrice", formula);

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
