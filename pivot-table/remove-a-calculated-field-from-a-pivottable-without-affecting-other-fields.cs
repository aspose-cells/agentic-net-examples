using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class RemoveCalculatedFieldDemo
    {
        // Entry point for the application
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string inputFile = "PivotTableWithCalculatedField.xlsx";
                const string outputFile = "PivotTable_CalculatedFieldRemoved.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook containing the pivot table with a calculated field
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the first pivot table in the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                PivotTable pivotTable = worksheet.PivotTables[0];

                // Name of the calculated field to be removed (must match the name used when adding it)
                string calculatedFieldName = "Profit";

                // Remove the calculated field from the Data area.
                // This operation does not affect other fields in the pivot table.
                pivotTable.RemoveField(PivotFieldType.Data, calculatedFieldName);

                // Recalculate the pivot table to reflect the removal
                pivotTable.CalculateData();

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}