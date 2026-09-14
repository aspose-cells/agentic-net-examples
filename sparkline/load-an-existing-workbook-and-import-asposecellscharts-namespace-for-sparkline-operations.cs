// Title: Insert a line sparkline into cell F1 of an existing Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load input.xlsx, add a line sparkline group that reads data from A1:D1, place it in cell F1, enable high‑point and low‑point markers, then save as output.xlsx with Aspose.Cells in C#. | Create a SparklineGroup of type Line on the first worksheet, set ShowHighPoint and ShowLowPoint to true, and persist the modified workbook to a new file using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells how to add a line sparkline to a specific cell | example of adding sparkline group from range A1:D1 in Aspose.Cells | set high and low point markers for sparkline using Aspose.Cells .NET | load existing Excel file and insert sparkline with Aspose.Cells C# tutorial
// Tags: line sparkline insertion Aspose.Cells C# | sparkline group creation from cell range Aspose.Cells | assign sparkline to target cell Aspose.Cells | high‑low point markers sparkline Aspose.Cells | workbook load and save Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads input.xlsx, adds a line sparkline group based on A1:D1 placed in cell F1, enables high‑point and low‑point markers, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Path to the existing workbook
        string inputPath = "input.xlsx";

        // Load the workbook from file
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define where the sparkline will be placed (cell F1)
        CellArea location = CellArea.CreateCellArea("F1", "F1");

        // Add a line sparkline group using data from A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Optional: customize the sparkline group appearance
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
