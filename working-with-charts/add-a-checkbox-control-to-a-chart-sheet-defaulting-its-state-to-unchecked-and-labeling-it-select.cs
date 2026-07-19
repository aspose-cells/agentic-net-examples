// Title: Add an Unchecked 'Select' Checkbox to an Excel Chart Sheet with Aspose.Cells (C#)
// Description: Demonstrates how to create a new workbook, add a chart sheet with a column chart, place a CheckBox control at row 5/column 5, set its label to "Select", keep it unchecked by default, and save the file as ChartSheetWithCheckBox.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart sheet checkbox | C# add checkbox to Excel chart | unchecked checkbox Aspose.Cells | Excel chart sheet control | Aspose.Cells C# example
// Common Searches: Aspose.Cells add checkbox to chart sheet C# | Create unchecked checkbox on Excel chart sheet | Set checkbox label to Select in Aspose.Cells | How to place a control on a chart sheet using Aspose.Cells | Default state of checkbox in Aspose.Cells workbook
// Developer Intent: Insert an unchecked checkbox labeled "Select" onto a chart sheet in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Build interactive dashboards where a checkbox toggles chart series visibility. | Generate reports that require user confirmation before processing chart data. | Automate workbook templates that include UI controls for downstream VBA or macro actions.
// AI Prompts: Write C# code with Aspose.Cells to add a checkbox to a chart sheet, set its text to "Select", and keep it unchecked. | Explain how to position a checkbox on a chart sheet and link its value to a cell using Aspose.Cells. | Show how to adjust the size and location of a checkbox on an Excel chart sheet in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new workbook, add a chart sheet with a column chart, place a CheckBox control at row 5/column 5, set its label to "Select", keep it unchecked by default, and save the file as ChartSheetWithCheckBox.xlsx using Aspose.Cells for .NET.
class AddCheckBoxToChartSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet that will serve as a chart sheet
            int chartSheetIndex = workbook.Worksheets.Add();
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "ChartSheet";

            // Add a simple column chart to the chart sheet (position and size in pixels)
            // Charts.Add returns the index of the newly created chart
            int chartIndex = chartSheet.Charts.Add(ChartType.Column, 0, 0, 400, 300);
            Chart chart = chartSheet.Charts[chartIndex];

            // Populate some data for the chart
            chartSheet.Cells["A1"].PutValue("Category");
            chartSheet.Cells["B1"].PutValue("Value");
            chartSheet.Cells["A2"].PutValue("Item 1");
            chartSheet.Cells["B2"].PutValue(10);
            chartSheet.Cells["A3"].PutValue("Item 2");
            chartSheet.Cells["B3"].PutValue(20);
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Add a checkbox to the chart sheet
            // Parameters: upperLeftRow, upperLeftColumn, height (pixels), width (pixels)
            int checkBoxIndex = chartSheet.CheckBoxes.Add(5, 5, 20, 100);
            CheckBox checkBox = chartSheet.CheckBoxes[checkBoxIndex];

            // Set the label text and ensure it is unchecked by default
            checkBox.Text = "Select";
            checkBox.Value = false; // unchecked

            // Save the workbook
            workbook.Save("ChartSheetWithCheckBox.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
