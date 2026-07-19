// Title: Aspose.Cells .NET: Create a Column Chart on a Hidden Sheet with Cell‑Linked Labels and Unhide the Sheet
// Description: Demonstrates how to hide a worksheet, populate it with categories, values and label text, add a column chart, enable cell‑based data labels linked to a separate column, set PlotVisibleCellsOnly to false, reveal the hidden sheet, and save the workbook as an XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart hidden worksheet | C# column chart cell linked labels | Aspose.Cells PlotVisibleCellsOnly | unhide worksheet Aspose.Cells | Aspose.Cells data labels from cells | create chart on hidden sheet | VisibilityType.Hidden Aspose.Cells | Excel automation hidden sheet chart
// Common Searches: Aspose.Cells add chart to hidden sheet | cell based data labels Aspose.Cells .NET | how to unhide worksheet after creating chart Aspose.Cells | PlotVisibleCellsOnly false hidden worksheet | link chart data labels to cells Aspose.Cells
// Developer Intent: Generate a column chart on a hidden worksheet, link its data labels to a cell range, then make the worksheet visible and save the file with Aspose.Cells for .NET.
// Use Cases: Prepare raw chart data on a hidden sheet to keep the workbook clean, then reveal the sheet for end‑user presentation. | Build a reporting template that stores intermediate calculations on a hidden sheet, applies cell‑linked labels, and publishes the visible sheet as the final report. | Automate Excel generation where charts must be created without exposing source data until processing is complete, then unhide the sheet for review.
// AI Prompts: Write C# code with Aspose.Cells to add a column chart on a hidden worksheet, link its data labels to a separate column, and then unhide the worksheet before saving. | Explain the effect of PlotVisibleCellsOnly on chart rendering when the source worksheet is hidden in Aspose.Cells. | Provide steps to remove a temporary visible worksheet after the hidden chart has been made visible in an Aspose.Cells workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates how to hide a worksheet, populate it with categories, values and label text, add a column chart, enable cell‑based data labels linked to a separate column, set PlotVisibleCellsOnly to false, reveal the hidden sheet, and save the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // First worksheet will hold the hidden chart
            Worksheet hiddenWs = workbook.Worksheets[0];
            hiddenWs.Name = "HiddenSheet";

            // Add a second worksheet to keep at least one visible sheet in the workbook
            Worksheet visibleWs = workbook.Worksheets.Add("VisibleSheet");

            // Hide the worksheet that will contain the chart
            hiddenWs.VisibilityType = VisibilityType.Hidden;

            // Populate sample data (categories, values, and label text)
            hiddenWs.Cells["A1"].PutValue("Category");
            hiddenWs.Cells["B1"].PutValue("Value");
            hiddenWs.Cells["C1"].PutValue("Label");
            hiddenWs.Cells["A2"].PutValue("A");
            hiddenWs.Cells["B2"].PutValue(10);
            hiddenWs.Cells["C2"].PutValue("Ten");
            hiddenWs.Cells["A3"].PutValue("B");
            hiddenWs.Cells["B3"].PutValue(20);
            hiddenWs.Cells["C3"].PutValue("Twenty");
            hiddenWs.Cells["A4"].PutValue("C");
            hiddenWs.Cells["B4"].PutValue(30);
            hiddenWs.Cells["C4"].PutValue("Thirty");

            // Add a column chart on the hidden worksheet
            int chartIdx = hiddenWs.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = hiddenWs.Charts[chartIdx];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable cell‑based data labels
            var series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;          // optional: show the numeric value
            series.DataLabels.ShowCellRange = true;      // enable cell‑range labels
            series.DataLabels.LinkedSource = "C2:C4";    // link labels to the cells in column C

            // Ensure hidden cells are plotted
            chart.PlotVisibleCellsOnly = false;

            // Make the hidden worksheet visible again
            hiddenWs.VisibilityType = VisibilityType.Visible;

            // Optionally remove the temporary visible worksheet if not needed
            // workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);

            // Save the workbook
            workbook.Save("ChartOnHiddenSheet.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
