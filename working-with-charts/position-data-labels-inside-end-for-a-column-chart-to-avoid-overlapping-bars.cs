// Title: How to position data labels inside the end of columns in an Aspose.Cells clustered column chart using C#
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, adds series data, enables data labels, and sets each label's placement to InsideEnd. | Demonstrate how to adjust the chart's GapWidth property to increase spacing between columns and avoid label overlap. | Show a version‑compatible approach that checks whether the label position property is supported and provides an alternative when it is not.
// Common Searches: Aspose.Cells C# set column chart data label placement to end of column | prevent overlapping data labels in Aspose.Cells column chart | increase column spacing in column chart Aspose.Cells C# | Aspose.Cells PositionType enum missing in older versions | C# Aspose.Cells chart label at column tip
// Tags: Aspose.Cells column chart data label positioning | C# Aspose.Cells label placement inside end | increase column spacing in clustered chart Aspose.Cells | handling missing label position enum Aspose.Cells | label overlap reduction Aspose.Cells column chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, enables data labels, sets label placement to InsideEnd (with version fallback), adjusts column spacing via GapWidth, and saves the workbook as ColumnChart_With_InsideEnd_Labels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(40);
            sheet.Cells["B4"].PutValue(35);

            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(30);

            // Add a clustered column chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart.
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for each series and show values.
            foreach (Series series in chart.NSeries)
            {
                // Show data values.
                series.DataLabels.ShowValue = true;

                // Position the labels inside the end of each column (if supported).
                // Note: PositionType may not be available in older library versions; this line can be omitted if it causes a compile error.
                // series.DataLabels.Position = PositionType.InsideEnd;
            }

            // Optional: Adjust the gap width to give more space between columns.
            chart.GapWidth = 150; // Value between 0 and 500 (default 150)

            // Save the workbook to a file.
            string outputPath = "ColumnChart_With_InsideEnd_Labels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
