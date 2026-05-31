using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class ReplacePivotTableDemo
    {
        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook containing the existing pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Get the collection of pivot tables on the worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // If there is at least one pivot table, remove the first one
                if (pivotTables.Count > 0)
                {
                    PivotTable oldPivot = pivotTables[0];
                    pivotTables.Remove(oldPivot);
                }

                // Define the source data range for the new pivot table
                // Here we assume the source data is in cells A1:B4 of the same worksheet
                string sourceData = "A1:B4";

                // Add a new pivot table at cell D5 with the name "NewPivot"
                int newIndex = pivotTables.Add(sourceData, "D5", "NewPivot");

                // Retrieve the newly added pivot table
                PivotTable newPivot = pivotTables[newIndex];

                // Configure the pivot table fields (example: Row = first column, Data = second column)
                newPivot.AddFieldToArea(PivotFieldType.Row, 0);   // First column as Row field
                newPivot.AddFieldToArea(PivotFieldType.Data, 1);  // Second column as Data field

                // Refresh and calculate the pivot table to populate it
                newPivot.RefreshData();
                newPivot.CalculateData();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ReplacePivotTableDemo.Run();
        }
    }
}