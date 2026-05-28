using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet (first sheet) containing the data
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Determine the used range of the source data
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.MaxDisplayRange;
            string sourceData = $"={sourceSheet.Name}!{sourceRange.Address}";

            // Create a new worksheet that will host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Add a pivot table
            PivotTableCollection pivotTables = pivotSheet.PivotTables;
            int pivotIndex = pivotTables.Add(sourceData, "A1", "MyPivotTable");

            // Configure pivot table fields (example)
            PivotTable pivotTable = pivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // First column as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Second column as data field

            // Save the workbook with the new pivot table
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table created and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}