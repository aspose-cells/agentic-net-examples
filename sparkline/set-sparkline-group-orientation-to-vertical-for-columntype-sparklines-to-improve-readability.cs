// Title: Add a vertical column sparkline group to an Excel worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a column‑type sparkline group, configures the group to display vertically, and writes the workbook to a file using Aspose.Cells. | Show how to enable the IsVertical flag on a SparklineGroup in Aspose.Cells and persist the changes.
// Common Searches: Aspose.Cells C# how to create a column sparkline with vertical orientation | set sparkline group orientation to vertical in .NET workbook | C# example for adding vertical column sparklines using Aspose.Cells | vertical sparkline group placement in Excel with Aspose.Cells API | Aspose.Cells SparklineGroup IsVertical property usage
// Tags: Aspose.Cells SparklineGroup IsVertical flag | C# sparkline group cell placement | Excel workbook sparkline creation Aspose | set IsVertical property Aspose.Cells | vertical sparkline example .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills cells A1:A5 with numbers, adds a column‑type sparkline group at B1 with the IsVertical flag set to true, and saves the file as SparklineVerticalOrientation.xlsx.
class SetSparklineOrientationVertical
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in a column (A1:A5)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(3);
            sheet.Cells["A3"].PutValue(7);
            sheet.Cells["A4"].PutValue(2);
            sheet.Cells["A5"].PutValue(9);

            // Define the location cell where the first sparkline will be placed (B1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add a Column‑type sparkline group with vertical orientation (isVertical = true)
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Column,
                "A1:A5",
                true,
                location);

            // Retrieve the created group (optional, for further customization)
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Save the workbook
            string outputPath = "SparklineVerticalOrientation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
