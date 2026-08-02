// Title: Update PivotTable Source Range with Aspose.Cells for .NET (C#)
// Description: Loads an existing Excel workbook, accesses the first worksheet and its first PivotTable, redefines the PivotTable's source range using the ChangeDataSource method, refreshes and recalculates the pivot data, and saves the modified file. Demonstrates how to programmatically change a PivotTable's data source in C# with Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | ChangeDataSource | RefreshData | CalculateData | update pivot source | modify pivot data range | Excel automation | programmatic pivot table
// Common Searches: Aspose.Cells change pivot table source range C# | How to update PivotTable data source with Aspose.Cells | RefreshData after changing pivot source Aspose.Cells | C# code to modify PivotTable source in Excel | Aspose.Cells ChangeDataSource example
// Developer Intent: Programmatically redefine the source data of an existing PivotTable and refresh its calculations using Aspose.Cells for .NET.
// Use Cases: Switch a PivotTable to a new data block after adding or removing columns. | Point multiple PivotTables to a consolidated data range without recreating them. | Adapt a PivotTable when the underlying dataset expands or contracts dynamically.
// AI Prompts: Generate C# code that updates a PivotTable to use a named range as its source with Aspose.Cells. | Show how to change the source data for all PivotTables in a workbook and refresh them using Aspose.Cells for .NET. | Explain how to keep PivotTable formatting intact while changing its data source programmatically.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an existing Excel workbook, accesses the first worksheet and its first PivotTable, redefines the PivotTable's source range using the ChangeDataSource method, refreshes and recalculates the pivot data, and saves the modified file. Demonstrates how to programmatically change a PivotTable's data source in C# with Aspose.Cells.
class UpdatePivotSource
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the first worksheet.");
                return;
            }

            // Retrieve the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Define the new data source range (e.g., C1:D10 on the same sheet)
            string newDataSource = $"{worksheet.Name}!C1:D10";

            // Change the pivot table's data source (expects a string array)
            pivotTable.ChangeDataSource(new string[] { newDataSource });

            // Refresh and recalculate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
