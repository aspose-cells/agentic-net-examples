// Title: Adjust Bar Chart Gap Width to 150 % with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add sample data, insert a column chart, set its GapWidth property to 150 % for tighter column spacing, and save the result as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart gap width | column chart spacing | GapWidth property | 150 percent bar chart | Excel chart formatting .NET | tight column spacing | chart customization Aspose | programmatic Excel chart
// Common Searches: Aspose.Cells set chart gap width | C# change column chart spacing | GapWidth 150 Aspose.Cells example | how to tighten bar chart columns in Excel code | chart.GapWidth valid range Aspose
// Developer Intent: Set the GapWidth of a column (bar) chart to 150 % to make columns appear closer together.
// Use Cases: Design a sales dashboard where dense column charts improve readability. | Generate financial reports that follow corporate style guidelines for chart spacing. | Create automated Excel exports with uniformly tightened bar charts across multiple worksheets.
// AI Prompts: Write C# code with Aspose.Cells that sets a column chart's GapWidth to 150 % and saves the workbook. | Explain the GapWidth property, its effect on column/bar charts, and the permissible value range in Aspose.Cells. | Provide a sample that creates several charts and assigns different GapWidth values to each for visual comparison.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to create a workbook, add sample data, insert a column chart, set its GapWidth property to 150 % for tighter column spacing, and save the result as an XLSX file using Aspose.Cells for .NET.
public class AdjustBarChartGapWidth
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Insert a column (bar) chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the gap width to 150% for tighter column spacing (valid range 0‑500)
        chart.GapWidth = 150;

        // Determine output file path
        string outputPath = "BarChartGapWidthAdjusted.xlsx";

        // Ensure the directory exists
        string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Save the workbook to a file
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
