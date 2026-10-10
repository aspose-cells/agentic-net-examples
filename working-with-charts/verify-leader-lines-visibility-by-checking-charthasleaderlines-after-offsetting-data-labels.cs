// Title: How to enable and confirm leader lines for a column chart series after moving data labels to OutsideEnd using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, sets the data label position to OutsideEnd, enables leader lines for the series, and prints the HasLeaderLines flag. | Show how to read the HasLeaderLines property of a chart series after adjusting label positions to verify leader line visibility in Aspose.Cells.
// Common Searches: Aspose.Cells C# enable leader lines on column chart series | Check HasLeaderLines property after setting data label position in Aspose.Cells | How to verify chart series leader lines visibility in .NET | Move data labels to OutsideEnd and show leader lines with Aspose.Cells | Read chart series leader line flag after label offset using Aspose.Cells for .NET
// Tags: Aspose.Cells chart series leader lines | C# column chart data label position OutsideEnd | Aspose.Cells HasLeaderLines property | verify chart leader line visibility .NET | enable leader lines Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, builds a column chart, shows data labels positioned OutsideEnd, enables leader lines for the series, and outputs the HasLeaderLines value to confirm visibility.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add series and set category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Show data labels and set their position
            chart.NSeries[0].DataLabels.ShowValue = true;               // corrected property
            chart.NSeries[0].DataLabels.Position = LabelPositionType.OutsideEnd;

            // Enable leader lines for the series
            chart.NSeries[0].HasLeaderLines = true;

            // Verify leader lines visibility
            bool leaderLinesVisible = chart.NSeries[0].HasLeaderLines;
            Console.WriteLine("Leader lines visible: " + leaderLinesVisible);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
