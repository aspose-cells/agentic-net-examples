// Title: C# – Add a Column Chart with Cell‑Based Labels on a Hidden Worksheet using Aspose.Cells
// Description: Demonstrates how to create a workbook, populate a hidden sheet with category, value and label data, insert a column chart on that sheet, link the chart’s data labels to a cell range, format the labels, hide the source worksheet, and save the file as XLSX.
// Keywords: Aspose.Cells hidden worksheet chart | C# column chart cell based labels | link data labels to cells Aspose | hide sheet after chart creation | Aspose.Cells PlotVisibleCellsOnly false
// Common Searches: Aspose.Cells create chart on hidden sheet C# | cell based data labels hidden worksheet Aspose | hide worksheet after adding chart .NET | link chart labels to cell range Aspose.Cells | plot data from hidden sheet Aspose
// Developer Intent: Generate a column chart on a hidden sheet, attach labels from a cell range, hide the data sheet, and export the workbook.
// Use Cases: Present confidential data through a chart while keeping the source sheet invisible. | Build a dashboard where only the chart is shown, with labels derived from hidden cells. | Automate report generation that requires cell‑based labels without exposing raw data.
// AI Prompts: Provide C# code that creates a line chart on a hidden worksheet, uses cells for data labels, and saves the workbook with the sheet hidden. | Show an example of a pie chart on a hidden sheet, enables cell‑based labels, customizes the label font, and hides the data worksheet before saving.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartOnHiddenSheet
{
    // Demonstrates how to create a workbook, populate a hidden sheet with category, value and label data, insert a column chart on that sheet, link the chart’s data labels to a cell range, format the labels, hide the source worksheet, and save the file as XLSX.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet.
                Workbook workbook = new Workbook();

                // The first worksheet will hold the hidden data.
                Worksheet hiddenSheet = workbook.Worksheets[0];
                hiddenSheet.Name = "HiddenData";

                // Add a second (dummy) worksheet to ensure the workbook always has at least one visible sheet.
                Worksheet dummySheet = workbook.Worksheets.Add("Dummy");

                // Populate sample data for the chart on the hidden sheet.
                hiddenSheet.Cells["A1"].PutValue("Category");
                hiddenSheet.Cells["B1"].PutValue("Value");
                hiddenSheet.Cells["C1"].PutValue("Label"); // Cell‑based labels

                hiddenSheet.Cells["A2"].PutValue("A");
                hiddenSheet.Cells["B2"].PutValue(10);
                hiddenSheet.Cells["C2"].PutValue("Ten");

                hiddenSheet.Cells["A3"].PutValue("B");
                hiddenSheet.Cells["B3"].PutValue(20);
                hiddenSheet.Cells["C3"].PutValue("Twenty");

                hiddenSheet.Cells["A4"].PutValue("C");
                hiddenSheet.Cells["B4"].PutValue(30);
                hiddenSheet.Cells["C4"].PutValue("Thirty");

                // Add a column chart to the hidden worksheet (the sheet can be visible while creating the chart).
                int chartIndex = hiddenSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = hiddenSheet.Charts[chartIndex];

                // Set the data range for the chart (values) and categories (X‑axis).
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable cell‑based data labels using the linked source cells.
                chart.NSeries[0].DataLabels.ShowCellRange = true;
                chart.NSeries[0].DataLabels.LinkedSource = "C2:C4";
                chart.NSeries[0].DataLabels.Font.Color = Color.Blue;
                chart.NSeries[0].DataLabels.Font.Size = 10;

                // Ensure hidden cells are plotted if needed.
                chart.PlotVisibleCellsOnly = false;

                // Hide the data worksheet now that the chart is created.
                hiddenSheet.VisibilityType = VisibilityType.Hidden;

                // Save the workbook.
                workbook.Save("ChartOnHiddenSheet.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }
    }
}
