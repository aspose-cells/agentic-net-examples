// Title: How to set a fixed vertical axis range (min 0, max 10) for a line sparkline using Aspose.Cells for .NET (C#)
// AI Prompts: Create an Excel workbook in C# with Aspose.Cells, add a line sparkline based on cells A1:D1, and set the sparkline group's vertical axis minimum to 0 and maximum to 10. | Using Aspose.Cells for .NET, configure a sparkline group's VerticalAxisMinValueType and VerticalAxisMaxValueType to Custom and assign specific numeric limits for consistent scaling. | Generate a .xlsx file that contains a line sparkline with an orange series color and fixed axis limits (0‑10) by applying custom axis settings in Aspose.Cells C#.
// Common Searches: Aspose.Cells C# set sparkline vertical axis minimum custom value | fixed axis scaling for line sparkline in .NET Excel using Aspose | how to enforce min and max values on sparkline group with Aspose.Cells | C# example of custom sparkline axis range Aspose.Cells | Aspose.Cells sparkline vertical axis custom range 0 to 10
// Tags: Aspose.Cells sparkline custom vertical axis | C# line sparkline fixed scaling | Aspose.Cells set sparkline axis min max | Excel sparkline axis limits Aspose.Cells | sparkline group custom axis values .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace SparklineAxisFixedScaling
{
    // The example creates a new workbook, fills cells A1:D1 with data, adds a line sparkline in cell E1, sets the sparkline group's vertical axis minimum to 0 and maximum to 10 using custom axis types, applies an orange series color, and saves the file as SparklineFixedAxisScaling.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group with a line sparkline, using the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (optional, already added by Add method)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Configure the vertical axis to use fixed minimum and maximum values
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Custom; // Use custom min
            group.VerticalAxisMinValue = 0.0;                                 // Fixed minimum

            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Custom; // Use custom max
            group.VerticalAxisMaxValue = 10.0;                                // Fixed maximum

            // Optional: set some visual properties for better visibility
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            // Save the workbook to a file
            workbook.Save("SparklineFixedAxisScaling.xlsx");
        }
    }
}
