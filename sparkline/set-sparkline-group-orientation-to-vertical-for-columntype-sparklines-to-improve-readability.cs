// Title: Add a Vertical Column Sparkline Group with Aspose.Cells for .NET (C#)
// Description: C# code that creates a workbook, fills A1:A5 with numbers, inserts a column‑type sparkline group plotted by column (vertical orientation) at B1, applies a preset style, and saves the file as SparklineVerticalOrientation.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# sparkline | vertical sparkline | column sparkline group | set sparkline orientation | Aspose.Cells for .NET example | sparkline preset style | Excel sparkline API | plot by column | vertical column sparkline code
// Common Searches: Aspose.Cells set sparkline orientation vertical | C# create vertical column sparkline | how to add column sparkline group in Aspose.Cells | vertical sparkline example .NET | plot sparkline by column Aspose.Cells
// Developer Intent: Create a column‑type sparkline group with vertical (plot‑by‑column) orientation in a worksheet using Aspose.Cells for .NET.
// Use Cases: Show monthly sales trends in a dashboard with a single‑cell vertical column sparkline for quick visual comparison. | Add a vertical sparkline next to KPI values so the trend is displayed by column, improving readability in financial reports. | Apply a consistent preset style to multiple vertical column sparklines across a workbook to maintain a unified look.
// AI Prompts: Generate C# code that uses Aspose.Cells to add a vertical column sparkline group from range D2:D12 and saves the workbook. | Show how to modify the line weight and color of an existing vertical column sparkline group with Aspose.Cells for .NET. | Provide an example that creates three vertical column sparkline groups on different rows, each with a distinct preset style, using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that creates a workbook, fills A1:A5 with numbers, inserts a column‑type sparkline group plotted by column (vertical orientation) at B1, applies a preset style, and saves the file as SparklineVerticalOrientation.xlsx using Aspose.Cells for .NET.
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

            // Define the location cell where the sparkline will be placed (B1)
            CellArea location = new CellArea
            {
                StartRow = 0,   // Row 1 (zero‑based)
                EndRow = 0,     // Single cell
                StartColumn = 1, // Column B
                EndColumn = 1
            };

            // Add a Column‑type sparkline group with vertical orientation.
            // isVertical = true means the sparkline is plotted by column.
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Column,   // Column sparkline type
                "A1:A5",                // Data range
                true,                   // Plot by column (vertical orientation)
                location);              // Where the sparkline appears

            // Retrieve the created group (optional, for further customization)
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Example: set a preset style (optional)
            group.PresetStyle = SparklinePresetStyleType.Style1;

            // Define output file path
            string outputPath = "SparklineVerticalOrientation.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
