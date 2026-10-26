// Title: Set PivotTable data fields to use the Average consolidation function with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change pivot table data field aggregation to average | How to apply average consolidation function to existing PivotTable using Aspose.Cells .NET | Update Excel PivotTable data fields to average values programmatically with Aspose.Cells | C# code to set PivotTable data field function to Average in Aspose.Cells | Modify PivotTable aggregation method to average in a workbook with Aspose.Cells for .NET
// Tags: set pivot table data field average Aspose.Cells | pivot table consolidation function average .NET | modify pivot table aggregation C# | Aspose.Cells change data field function | average aggregation in Excel pivot using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The program loads an existing Excel workbook, verifies that the first worksheet contains a PivotTable, iterates over each data field of that PivotTable, sets the field's Function property to ConsolidationFunction.Average, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found on the first worksheet.");
                return;
            }

            // Get the first PivotTable on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Change the consolidation function of each data field to Average
            foreach (PivotField dataField in pivotTable.DataFields)
            {
                dataField.Function = ConsolidationFunction.Average;
            }

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
