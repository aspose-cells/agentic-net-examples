// Title: Aspose.Cells C# – Hide Tick Marks on Secondary Y‑Axis of a Mixed Column Chart
// Description: This example creates an Excel workbook, adds sample data, builds a mixed column chart with primary and secondary series, assigns the second series to the secondary axis, and disables both major and minor tick marks on that axis before saving the file.
// Keywords: Aspose.Cells | C# | mixed column chart | secondary Y axis | tick mark none | chart axis formatting | Excel chart customization | TickMarkType.None
// Common Searches: Aspose.Cells hide secondary axis tick marks | C# remove tick marks from secondary Y axis in Excel chart | set secondary value axis tick marks to none Aspose | mixed chart secondary axis formatting .NET | disable minor tick marks secondary axis Aspose.Cells
// Developer Intent: Remove major and minor tick marks from the secondary Y‑axis of a mixed chart.
// Use Cases: Generate a clean dashboard by suppressing secondary axis tick marks in Excel charts | Programmatically customize chart axes for reports using Aspose.Cells | Create mixed charts where the secondary axis shows only labels without tick marks | Automate Excel chart styling in .NET applications to improve readability
// AI Prompts: Provide C# Aspose.Cells code that creates a mixed column chart and sets TickMarkType.None for both major and minor tick marks on the secondary Y‑axis. | Show how to assign a data series to the secondary axis and hide its tick marks using Aspose.Cells for .NET. | Explain step‑by‑step how to customize secondary axis properties such as tick marks, line style, and label format in an Excel workbook with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates an Excel workbook, adds sample data, builds a mixed column chart with primary and secondary series, assigns the second series to the secondary axis, and disables both major and minor tick marks on that axis before saving the file.
class HideSecondaryAxisTickMarks
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");

        sheet.Cells["B1"].PutValue("Primary");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Secondary");
        sheet.Cells["C2"].PutValue(500);
        sheet.Cells["C3"].PutValue(600);
        sheet.Cells["C4"].PutValue(700);

        // Add a mixed column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: primary and secondary
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.Add("C2:C4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Plot the second series on the secondary Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Hide tick marks on the secondary Y axis
        Axis secondaryYAxis = chart.SecondValueAxis;
        secondaryYAxis.MajorTickMark = TickMarkType.None;
        secondaryYAxis.MinorTickMark = TickMarkType.None;

        // Save the workbook
        workbook.Save("MixedChart_SecondaryAxis_NoTickMarks.xlsx");
    }
}
