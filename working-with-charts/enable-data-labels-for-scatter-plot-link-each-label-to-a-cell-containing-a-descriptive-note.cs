// Title: How to add cell‑linked data labels to a scatter chart in Aspose.Cells for C#
// Common Searches: Aspose.Cells C# scatter chart data labels linked to worksheet cells | set LinkedSource property for chart data labels in Aspose.Cells | display custom notes as data labels on a scatter plot using Aspose.Cells | C# example to show cell values as labels on scatter chart points with Aspose.Cells | configure data label position and format for Aspose.Cells scatter chart
// Tags: linked data labels for scatter charts Aspose.Cells | Aspose.Cells LinkedSource usage | custom chart point notes in .NET | data label formatting options Aspose.Cells | export scatter chart with linked labels to XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills columns A, B, and C with X values, Y values, and descriptive notes, adds a scatter chart, enables data labels, links each label to the corresponding note in column C via the LinkedSource property, customizes label position and appearance, and saves the file as ScatterChartWithLinkedLabels.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Populate sample data:
        // Column A -> X values
        // Column B -> Y values
        // Column C -> Descriptive notes for each point
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("X");
        sheet.Cells["B1"].PutValue("Y");
        sheet.Cells["C1"].PutValue("Note");

        // Example data (4 points)
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue("First point");

        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["C3"].PutValue("Second point");

        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue(15);
        sheet.Cells["C4"].PutValue("Third point");

        sheet.Cells["A5"].PutValue(4);
        sheet.Cells["B5"].PutValue(25);
        sheet.Cells["C5"].PutValue("Fourth point");

        // -------------------------------------------------
        // Add a scatter chart
        // -------------------------------------------------
        int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set Y values range for the series
        chart.NSeries.Add("B2:B5", true);
        // Set X values range for the series
        chart.NSeries[0].XValues = "A2:A5";

        // -------------------------------------------------
        // Enable data labels and link each label to the note in column C
        // -------------------------------------------------
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;               // Show the Y value
        dataLabels.ShowCellRange = true;           // Enable linking to a cell range
        dataLabels.LinkedSource = "C2:C5";         // Range that contains the notes

        // Optional: adjust label appearance
        dataLabels.Position = LabelPositionType.Above; // Position label above the point
        dataLabels.IsAutoText = false;                 // Use the linked text only
        dataLabels.NumberFormatLinked = false;         // Do not inherit number format from cells

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        try
        {
            string outputPath = "ScatterChartWithLinkedLabels.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error while saving the workbook: " + ex.Message);
        }
    }
}
