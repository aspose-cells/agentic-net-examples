// Title: Increase the DepthPercent of a Column3D chart in Aspose.Cells for .NET to enhance 3‑D visual depth
// AI Prompts: Generate C# sample that builds a workbook, inserts a Column3D chart, and assigns a larger DepthPercent (e.g., 300) using Aspose.Cells. | Show how to combine DepthPercent with Perspective and Elevation settings to create a deeper 3‑D effect for a chart in Aspose.Cells.
// Common Searches: how to set depthpercent for a 3d column chart in Aspose.Cells C# | increase visual depth of Aspose.Cells Column3D chart programmatically | adjust perspective and elevation together with depthpercent Aspose.Cells .NET
// Tags: Aspose.Cells Column3D chart DepthPercent | C# set 3d chart depth Aspose.Cells | chart visual depth adjustment Aspose.Cells | Aspose.Cells chart perspective elevation | DepthPercent range 20-2000 Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a 3‑D Column chart, and enhances its visual depth by setting the DepthPercent property to 300, with optional Perspective and Elevation tweaks, then saves the file as SparklineDepthDemo.xlsx.
class SetSparklineDepthDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a 3‑D chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);
        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(90);
        sheet.Cells["C3"].PutValue(110);
        sheet.Cells["C4"].PutValue(130);

        // Add a 3‑D column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:C4", true);          // Add series data
        chart.NSeries.CategoryData = "A2:A4";      // Set category axis

        // Increase visual depth perception by setting a higher DepthPercent value
        // Valid range is 20‑2000; 300 gives a noticeably deeper look
        chart.DepthPercent = 300;

        // Optional: tweak other 3‑D view properties for better effect
        chart.Perspective = 40;   // Perspective angle (0‑100)
        chart.Elevation = 30;     // Elevation angle (-90‑90)

        // Save the workbook
        workbook.Save("SparklineDepthDemo.xlsx", SaveFormat.Xlsx);
    }
}
