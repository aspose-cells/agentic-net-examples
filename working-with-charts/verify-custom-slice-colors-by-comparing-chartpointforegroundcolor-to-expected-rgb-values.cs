// Title: C# – Verify Pie Chart Slice Colors Using ChartPoint.Area.ForegroundColor in Aspose.Cells
// Description: Creates a workbook, adds a pie chart with three categories, assigns specific RGB colors to each slice via ChartPoint.Area.ForegroundColor, then checks each slice's actual color against the expected values and logs matches or mismatches before saving the file.
// Keywords: Aspose.Cells pie chart slice color | ChartPoint Area ForegroundColor C# | verify chart point RGB Aspose | custom pie slice colors .NET | Excel chart color validation | Aspose.Cells color comparison
// Common Searches: set custom colors for pie chart slices Aspose.Cells | compare ChartPoint.ForegroundColor with expected RGB | how to validate pie slice colors in C# Excel library | Aspose.Cells verify chart point colors programmatically | check pie chart slice color matches in .NET
// Developer Intent: Programmatically confirm that each pie chart slice is rendered with the intended RGB color by comparing the applied ForegroundColor to a predefined color array.
// Use Cases: Enforce brand color guidelines in automatically generated Excel reports. | Automated UI testing of Excel chart appearance for continuous integration pipelines. | Generate Excel dashboards with colored slices and validate visual accuracy before distribution.
// AI Prompts: Generate C# code that sets custom RGB colors for each slice of a pie chart using Aspose.Cells and then verifies the colors match a given list. | Provide a reusable method that takes a Chart object and a Color[] array, applies the colors to ChartPoint.Area.ForegroundColor, and returns any mismatched indices. | Explain how to retrieve ChartPoint.Area.ForegroundColor values and compare them to expected colors in an Aspose.Cells workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a pie chart with three categories, assigns specific RGB colors to each slice via ChartPoint.Area.ForegroundColor, then checks each slice's actual color against the expected values and logs matches or mismatches before saving the file.
class VerifyPieSliceColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data for the pie chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(45);
        sheet.Cells["B4"].PutValue(25);

        // Add a pie chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Define the expected RGB colors for each slice
        Color[] expectedColors = new Color[]
        {
            Color.FromArgb(255, 0, 0),   // Red for Apple
            Color.FromArgb(0, 255, 0),   // Green for Banana
            Color.FromArgb(0, 0, 255)    // Blue for Cherry
        };

        // Apply the custom colors to each chart point (slice)
        for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
        {
            ChartPoint point = chart.NSeries[0].Points[i];
            point.Area.ForegroundColor = expectedColors[i];
        }

        // Verify that the applied colors match the expected RGB values
        for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
        {
            ChartPoint point = chart.NSeries[0].Points[i];
            Color actual = point.Area.ForegroundColor;
            Color expected = expectedColors[i];
            bool isMatch = actual.ToArgb() == expected.ToArgb();

            Console.WriteLine(
                $"Slice {i + 1}: Expected RGB({expected.R},{expected.G},{expected.B}) " +
                $"- Actual RGB({actual.R},{actual.G},{actual.B}) => " +
                (isMatch ? "Match" : "Mismatch"));
        }

        // Save the workbook to verify the result
        workbook.Save("PieSliceColorVerification.xlsx");
    }
}
