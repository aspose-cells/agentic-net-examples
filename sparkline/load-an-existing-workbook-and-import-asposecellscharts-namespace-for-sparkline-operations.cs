using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for sparkline types and groups

class SparklineExample
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // Load the workbook using the string constructor (load rule)
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is some data for the sparkline (A1:D1)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(3);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(2);

        // Define where the sparkline will be placed (cell E1)
        CellArea location = CellArea.CreateCellArea("E1", "E1");

        // Add a line sparkline group with the specified data range and location
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,          // Sparkline type
            "A1:D1",                     // Data range
            false,                       // Plot by row (horizontal)
            location                     // Destination cell area
        );

        // Optional: customize the sparkline group (e.g., show high/low points)
        SparklineGroup group = sheet.SparklineGroups[groupIndex];
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the modified workbook (save rule)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);

        Console.WriteLine($"Workbook loaded from '{inputPath}', sparkline added, and saved to '{outputPath}'.");
    }
}