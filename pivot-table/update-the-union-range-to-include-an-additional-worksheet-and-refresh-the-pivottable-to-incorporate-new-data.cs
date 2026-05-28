using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUpdate
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Input and output file paths
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that should contain a PivotTable
                Workbook workbook = new Workbook(inputPath);

                // Assume original data is on the first worksheet
                Worksheet sourceSheet = workbook.Worksheets[0];
                Worksheet pivotSheet = sourceSheet; // will be reassigned if a new PivotTable is created

                // Add a new worksheet that will hold additional data for the union range
                Worksheet newDataSheet = workbook.Worksheets.Add("AdditionalData");

                // Populate the new worksheet with sample data (same layout as original)
                newDataSheet.Cells["A1"].PutValue("Product");
                newDataSheet.Cells["B1"].PutValue("Sales");
                newDataSheet.Cells["A2"].PutValue("Gadget");
                newDataSheet.Cells["B2"].PutValue(1500);
                newDataSheet.Cells["A3"].PutValue("Widget");
                newDataSheet.Cells["B3"].PutValue(2300);
                newDataSheet.Cells["A4"].PutValue("Thingamajig");
                newDataSheet.Cells["B4"].PutValue(1200);

                // Ensure that a PivotTable exists before attempting to modify it
                if (pivotSheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTable found in the workbook. Creating a new PivotTable.");

                    // Create a simple PivotTable based on the original data range
                    string sourceRange = $"{sourceSheet.Name}!A1:B4";
                    Worksheet ptSheet = workbook.Worksheets.Add("PivotSheet");
                    ptSheet.PivotTables.Add(sourceRange, "A3", "PivotTable1");

                    // Use the newly created sheet for subsequent operations
                    pivotSheet = ptSheet;
                }

                // Retrieve the first PivotTable
                PivotTable pivotTable = pivotSheet.PivotTables[0];

                // Build the union data source that includes both the original and the new worksheet ranges
                string[] unionSource = new string[]
                {
                    $"{sourceSheet.Name}!A1:B4",   // Original data range
                    $"{newDataSheet.Name}!A1:B4"   // New data range
                };

                // Change the PivotTable's data source to the union range (string[] overload)
                pivotTable.ChangeDataSource(unionSource);

                // Refresh the PivotTable to reflect the updated data source
                pivotSheet.RefreshPivotTables();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}