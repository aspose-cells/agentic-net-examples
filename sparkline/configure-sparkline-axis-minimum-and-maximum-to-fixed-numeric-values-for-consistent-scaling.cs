// Title: Set Fixed Minimum and Maximum for Sparkline Axis in Aspose.Cells (C#/.NET)
// Description: Creates a workbook, adds sample data, inserts a line sparkline in cell E1, and configures the sparkline group's vertical axis to custom limits (0.0 – 10.0) using SparklineAxisMinMaxType.Custom before saving the file.
// Keywords: Aspose.Cells | sparkline custom axis | C# | .NET | vertical axis min max | fixed scaling | SparklineAxisMinMaxType.Custom
// Common Searches: Aspose.Cells set sparkline axis minimum and maximum C# | fixed sparkline scaling Aspose.Cells .NET example | how to use SparklineAxisMinMaxType.Custom | C# sparkline vertical axis limits Aspose.Cells
// Developer Intent: Apply explicit numeric limits to a sparkline’s vertical axis for consistent visual scaling across worksheets.
// Use Cases: Compare multiple rows of data where each sparkline shares a 0‑10 scale. | Produce financial dashboards with uniform sparkline ranges to highlight trend differences. | Prepare reports that require identical axis settings before exporting to Excel.
// AI Prompts: Show how to set both vertical and horizontal custom axis limits for a sparkline group in Aspose.Cells (C#). | Provide a C# sample that creates several sparkline groups with the same fixed axis range using Aspose.Cells. | Explain the role of SparklineAxisMinMaxType.Custom and the required properties to define fixed axis values.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a line sparkline in cell E1, and configures the sparkline group's vertical axis to custom limits (0.0 – 10.0) using SparklineAxisMinMaxType.Custom before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data that the sparkline will represent
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(3);

        // Define where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // column E (0‑based index)
            EndColumn = 4
        };

        // Add a sparkline group using the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Configure the vertical axis to use fixed numeric limits
        group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Custom; // use custom minimum
        group.VerticalAxisMinValue = 0.0;                                 // fixed minimum value

        group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Custom; // use custom maximum
        group.VerticalAxisMaxValue = 10.0;                                // fixed maximum value

        // Save the workbook to a file
        workbook.Save("SparklineFixedAxis.xlsx");
    }
}
