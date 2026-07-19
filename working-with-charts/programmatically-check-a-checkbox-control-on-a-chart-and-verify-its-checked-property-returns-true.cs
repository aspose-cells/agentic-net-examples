// Title: Aspose.Cells for .NET: Set and Verify a CheckBox on a Chart after Saving and Reloading
// Description: Demonstrates how to add a CheckBox shape over a chart, programmatically set its Value to true, save the workbook, reload it, and confirm that the checkbox remains checked using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | checkbox on chart | set checkbox value | verify checkbox state | save and reload workbook | Excel automation | chart dashboard | Aspose.Cells example
// Common Searches: Aspose.Cells set checkbox value on chart | how to read checkbox state after saving Excel with Aspose.Cells | C# programmatically check a shape over a chart | verify checkbox is true after workbook reload | Aspose.Cells example for chart checkbox
// Developer Intent: Programmatically check a CheckBox placed on a chart, persist the workbook, reload it, and ensure the checkbox’s Value property is true.
// Use Cases: Create interactive Excel dashboards where checkboxes control chart visibility and must retain their state across saves. | Generate report templates with pre‑selected options that are validated when the file is opened. | Automate quality‑assurance tests that confirm UI controls embedded in charts persist correctly after a round‑trip.
// AI Prompts: Show C# code to add a CheckBox over a chart, set its Value to true, save the workbook, reload it, and verify the checked state with Aspose.Cells. | Provide error‑handling that throws an exception if a reloaded workbook does not contain a checked CheckBox. | Explain how to iterate through all CheckBox objects on a worksheet containing charts and output each checkbox’s checked status.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to add a CheckBox shape over a chart, programmatically set its Value to true, save the workbook, reload it, and confirm that the checkbox remains checked using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create ----------
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["B1"].PutValue(15);
            sheet.Cells["B2"].PutValue(25);
            sheet.Cells["B3"].PutValue(35);

            // Add a column chart that uses the data range
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("A1:B3", true);
            chart.Title.Text = "Sample Chart";

            // Add a CheckBox shape positioned over the chart area
            // Parameters: upper-left row, upper-left column, height (pixels), width (pixels)
            int checkBoxIdx = sheet.CheckBoxes.Add(6, 2, 20, 100);
            CheckBox checkBox = sheet.CheckBoxes[checkBoxIdx];
            checkBox.Text = "Check me";
            // Programmatically check the box
            checkBox.Value = true;

            // ---------- Save ----------
            string fileName = "CheckboxOnChart.xlsx";
            workbook.Save(fileName);

            // ---------- Load ----------
            if (!File.Exists(fileName))
                throw new FileNotFoundException($"The file '{fileName}' was not found after saving.");

            Workbook loadedWb = new Workbook(fileName);
            Worksheet loadedSheet = loadedWb.Worksheets[0];

            if (loadedSheet.CheckBoxes.Count == 0)
                throw new InvalidOperationException("No CheckBox objects were found in the loaded worksheet.");

            CheckBox loadedCheckBox = loadedSheet.CheckBoxes[0];

            // Verify that the checkbox is checked
            bool isChecked = loadedCheckBox.Value; // should be true
            Console.WriteLine("Checkbox Checked: " + isChecked);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
