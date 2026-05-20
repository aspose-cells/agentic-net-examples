using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputPath = "input.xlsx";

        // Load the workbook from the file
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Example sparkline operation:
        // Define where the sparkline will be placed (cell F1)
        CellArea location = CellArea.CreateCellArea("F1", "F1");

        // Add a line‑type sparkline group using data from A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optional customizations
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}