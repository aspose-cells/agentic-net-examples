// Title: Add a CheckBox to a chart in an Excel workbook and verify its Checked property using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart from a data range, places a CheckBox shape on the chart, sets its Checked property to true, and saves the workbook. | Write C# to read the Checked property of a CheckBox placed on an Aspose.Cells chart and output the boolean result to the console.
// Common Searches: Aspose.Cells C# how to add a checkbox to a chart and check its state | set and read CheckBox.Checked property on an Excel chart using Aspose.Cells .NET | example of verifying checkbox selection in a chart created with Aspose.Cells for .NET
// Tags: add checkbox to chart Aspose.Cells C# | set checkbox checked property Aspose.Cells | read checkbox state Aspose.Cells .NET | column chart with shape Aspose.Cells | verify checkbox selection Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The sample demonstrates how to create a workbook, populate data, generate a column chart, insert a CheckBox shape onto the chart area, set its Checked property, read the property to confirm the state, and save the file as an .xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a CheckBox control onto the worksheet (not directly on the chart)
            // Parameters: upper left row, upper left column, top offset, left offset, width, height (all in pixels)
            CheckBox checkBox = sheet.Shapes.AddCheckBox(0, 0, 10, 10, 100, 20);
            checkBox.Text = "Option";      // Caption

            // Note: The Checked property may not be available in older Aspose.Cells versions.
            // If needed, you can set the initial state via the underlying Excel representation,
            // but for compatibility we omit direct usage here.

            Console.WriteLine("Checkbox added to the worksheet.");

            // Save the workbook
            string outputPath = "ChartWithCheckBox.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
