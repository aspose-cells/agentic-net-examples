// Title: Aspose.Cells for .NET: Programmatically check a checkbox shape on a chart and confirm its Value
// Description: This C# example creates a workbook, adds a column chart, places a checkbox shape over the chart, sets the checkbox Value to true, reads the Value back to verify it is checked, and saves the file as an .xlsx workbook.
// Keywords: Aspose.Cells | .NET | C# | checkbox shape | chart control | set checkbox value | verify checked state | programmatic Excel | Workbook.Save | Excel automation
// Common Searches: how to set a checkbox on an Aspose.Cells chart programmatically | verify checkbox Value property in Asp.NET | add checkbox shape to Excel chart using Aspose.Cells | C# code to check a checkbox over a chart | Aspose.Cells example for checkbox Value true
// Developer Intent: Set a checkbox placed on a chart to checked and read its Value to ensure it returns true.
// Use Cases: Build interactive Excel dashboards where checkboxes over charts control visible data series. | Create report templates that require pre‑selected options before distribution. | Generate workbooks for downstream processing that rely on checkbox states to trigger logic.
// AI Prompts: Generate C# code with Aspose.Cells that adds a checkbox shape on a chart, marks it as checked, and prints the verification result. | Show how to programmatically read the Value of a checkbox placed on an Excel chart using Aspose.Cells for .NET. | Provide an Aspose.Cells example that saves a workbook containing a chart with a pre‑checked checkbox and outputs the checked status.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This C# example creates a workbook, adds a column chart, places a checkbox shape over the chart, sets the checkbox Value to true, reads the Value back to verify it is checked, and saves the file as an .xlsx workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(15);
            sheet.Cells["B2"].PutValue(25);
            sheet.Cells["B3"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("A1:A3", true);
            chart.NSeries[0].Name = "Series1";
            chart.NSeries.Add("B1:B3", true);
            chart.NSeries[1].Name = "Series2";

            // Add a checkbox shape positioned over the chart area
            // Parameters: upperLeftRow, upperLeftColumn, width, height
            int checkBoxIdx = sheet.CheckBoxes.Add(6, 1, 20, 100);
            CheckBox checkBox = sheet.CheckBoxes[checkBoxIdx];
            checkBox.Text = "Sample Checkbox";

            // Programmatically check the checkbox
            checkBox.Value = true; // Equivalent to CheckedValue = CheckValueType.Checked

            // Verify that the checkbox is checked
            bool isChecked = checkBox.Value;
            Console.WriteLine("Checkbox Checked? " + isChecked); // Should output True

            // Save the workbook
            string outputPath = "CheckboxOnChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
