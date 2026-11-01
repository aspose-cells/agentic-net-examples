// Title: Add a Profit calculated field ([Revenue]-[Cost]) to an existing Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: how to add a calculated field named Profit to a PivotTable with Aspose.Cells in C# | Aspose.Cells C# add custom formula [Revenue]-[Cost] to an existing PivotTable | programmatically insert a calculated field into an Excel PivotTable using Aspose.Cells .NET | C# code to create a Profit field in a PivotTable from Revenue and Cost columns | Aspose.Cells example for adding a calculated field to a workbook PivotTable
// Tags: calculated field addition PivotTable Aspose.Cells | Profit field expression Revenue minus Cost C# | modify existing PivotTable programmatically Aspose.Cells | Excel workbook calculated field automation .NET | Aspose.Cells custom PivotTable formula example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot; // Required for PivotTable class

// The sample loads an existing Excel workbook, checks for a PivotTable on the first worksheet, adds a calculated field called "Profit" with the formula "[Revenue]-[Cost]", and saves the updated file.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet where the PivotTable resides (first worksheet in this example)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No PivotTables found on the first worksheet.");
                return;
            }

            // Retrieve the first PivotTable on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Add a calculated field named "Profit" with the expression [Revenue]-[Cost]
            pivotTable.AddCalculatedField("Profit", "[Revenue]-[Cost]");

            // Save the workbook with the new calculated field
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
