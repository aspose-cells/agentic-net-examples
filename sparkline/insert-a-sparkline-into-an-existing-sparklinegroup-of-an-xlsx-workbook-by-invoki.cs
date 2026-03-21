using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class InsertSparklineIntoExistingGroup
{
    static void Main()
    {
        // Load an existing XLSX workbook that already contains a SparklineGroup
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the collection of SparklineGroups in the worksheet
        SparklineGroupCollection groups = sheet.SparklineGroups;

        // Ensure there is at least one SparklineGroup to work with
        if (groups.Count == 0)
        {
            Console.WriteLine("No SparklineGroup found in the worksheet.");
            return;
        }

        // Retrieve the first existing SparklineGroup (or any specific index you need)
        SparklineGroup existingGroup = groups[0];

        // Define the data range for the new sparkline (e.g., cells A1:D1)
        string dataRange = sheet.Name + "!A1:D1";

        // Define the location where the sparkline will be placed (row and column indexes)
        // Here we place it at row 2, column 5 (zero‑based indexes)
        int targetRow = 2;
        int targetColumn = 5;

        // Add the sparkline to the existing group using the SparklineCollection.Add method
        int sparklineIndex = existingGroup.Sparklines.Add(dataRange, targetRow, targetColumn);

        // Optionally, you can retrieve the newly added Sparkline object for further customization
        Sparkline newSparkline = existingGroup.Sparklines[sparklineIndex];
        Console.WriteLine($"Added sparkline at row {newSparkline.Row}, column {newSparkline.Column} with data range {newSparkline.DataRange}.");

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}