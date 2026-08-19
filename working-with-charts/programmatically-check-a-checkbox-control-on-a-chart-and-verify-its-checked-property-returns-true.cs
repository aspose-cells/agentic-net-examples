// Title: Programmatically check a checkbox on an Excel chart and validate its state with Aspose.Cells for .NET
// Description: This example creates a workbook, adds a column chart, overlays a CheckBox shape on the chart, sets the CheckBox.Value to true, saves the file, reloads it, and reads the Value property to confirm the checkbox remains checked.
// Keywords: Aspose.Cells checkbox chart | C# set checkbox checked | Excel checkbox Value property | Aspose.Cells CheckBox API | persist checkbox state | chart overlay control Aspose.Cells | load workbook checkbox state | Aspose.Cells for .NET interactive controls
// Common Searches: how to set a checkbox as checked on an Excel chart using Aspose.Cells | read checkbox Value after saving workbook with Aspose.Cells | Aspose.Cells add checkbox over chart and verify state | C# Aspose.Cells checkboxes on charts | verify checkbox remains checked after workbook reload
// Developer Intent: The developer needs to programmatically check a checkbox placed on a chart and ensure its Value property returns true after the workbook is saved and reopened.
// Use Cases: Add a checked checkbox on a chart to let end‑users toggle features directly in the generated Excel file. | Automate quality checks that confirm interactive controls retain their state after file export. | Create unit tests that generate a chart with a checked checkbox, persist the workbook, and assert the checkbox stays checked.
// AI Prompts: Generate C# code with Aspose.Cells that adds a checkbox over a chart, sets it to checked, saves the workbook, reloads it, and verifies the checkbox Value is true. | Write a C# unit test using Aspose.Cells to create a column chart with a checked checkbox, persist the file, load it back, and assert the checkbox remains checked. | Explain how to access the Value property of a CheckBox placed on an Excel chart with Aspose.Cells, including handling cases where no checkboxes exist.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsCheckboxOnChart
{
    // This example creates a workbook, adds a column chart, overlays a CheckBox shape on the chart, sets the CheckBox.Value to true, saves the file, reloads it, and reads the Value property to confirm the checkbox remains checked.
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
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["A3"].PutValue("Item 2");
                sheet.Cells["A4"].PutValue("Item 3");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart that uses the data range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Add a checkbox shape positioned over the chart area
                // Parameters: upper left row, upper left column, upper left pixel offset, lower right pixel offset
                int checkBoxIndex = sheet.CheckBoxes.Add(7, 2, 20, 100);
                CheckBox checkBox = sheet.CheckBoxes[checkBoxIndex];

                // Set checkbox properties and check it
                checkBox.Text = "Enable Feature";
                checkBox.Value = true; // Checked state

                // Save the workbook to a file
                string filePath = "CheckboxOnChart.xlsx";
                workbook.Save(filePath);

                // Verify the file exists before loading
                if (File.Exists(filePath))
                {
                    // Load the workbook again to verify the checkbox state
                    Workbook loadedWorkbook = new Workbook(filePath);
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                    if (loadedSheet.CheckBoxes.Count > 0)
                    {
                        CheckBox loadedCheckBox = loadedSheet.CheckBoxes[0];
                        bool isChecked = loadedCheckBox.Value; // Returns true if checked
                        Console.WriteLine("Checkbox checked state after reload: " + isChecked);
                    }
                    else
                    {
                        Console.WriteLine("No checkboxes found in the loaded workbook.");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to create the workbook file: " + filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
