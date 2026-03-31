using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Access the collection of sparkline groups on the worksheet
        SparklineGroupCollection groups = sheet.SparklineGroups;

        // Ensure there is at least one sparkline group; create one if none exist
        if (groups.Count == 0)
        {
            // Add a new sparkline group of type Line (no initial sparklines)
            groups.Add(SparklineType.Line);
        }

        // Retrieve the first (or existing) sparkline group
        SparklineGroup group = groups[0];

        // Define the data range for the new sparkline (e.g., cells A2:D2 on the same sheet)
        string dataRange = sheet.Name + "!A2:D2";

        // Define the location where the sparkline will be placed (row 1, column 5 -> zero‑based indices)
        int targetRow = 1;      // corresponds to Excel row 2
        int targetColumn = 5;   // corresponds to Excel column F

        // Insert the new sparkline into the group's Sparklines collection
        int sparklineIndex = group.Sparklines.Add(dataRange, targetRow, targetColumn);

        // Optional: customize the sparkline group appearance
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}