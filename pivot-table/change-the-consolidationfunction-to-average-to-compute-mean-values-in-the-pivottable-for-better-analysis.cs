// Title: Change PivotTable data field aggregation to Average using Aspose.Cells for .NET (C#)
// Common Searches: how to set pivot table data field aggregation to average in C# with Aspose.Cells | Aspose.Cells set ConsolidationFunction.Average for Excel pivot tables | C# example of updating pivot table calculations to average using Aspose.Cells | modify pivot table data fields to average consolidation function in .NET | sample code to change pivot table aggregation method to average with Aspose.Cells
// Tags: Aspose.Cells pivot table aggregation average | C# set ConsolidationFunction.Average | update Excel pivot data fields Aspose.Cells | modify pivot table calculations .NET | change pivot table data field function C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for worksheets and a pivot table, iterates over the pivot table's data fields, sets each field's ConsolidationFunction to Average, and saves the workbook with the updated aggregation settings.
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

                // Ensure the workbook contains at least one worksheet
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
                    Console.WriteLine("No pivot tables found in the selected worksheet.");
                    return;
                }

                // Retrieve the first pivot table on the worksheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Set the consolidation function of each data field to Average
                foreach (PivotField dataField in pivotTable.DataFields)
                {
                    dataField.Function = ConsolidationFunction.Average;
                }

                // Save the workbook with the updated pivot table
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
