using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data that will be used for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Create a column chart placed near the top of the sheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 7);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Add a checkbox shape that sits on top of the chart area
            // Parameters: upper left row, upper left column, width (pixels), height (pixels)
            int checkBoxIdx = sheet.CheckBoxes.Add(6, 2, 100, 20);
            CheckBox checkBox = sheet.CheckBoxes[checkBoxIdx];
            checkBox.Text = "Check me";

            // Programmatically set the checkbox to the checked state
            checkBox.Value = true; // Boolean property
            checkBox.CheckedValue = CheckValueType.Checked; // Enum property (optional)

            // Verify that the checkbox is indeed checked
            Console.WriteLine("Checkbox.Value (expected True): " + checkBox.Value);
            Console.WriteLine("Checkbox.CheckedValue (expected Checked): " + checkBox.CheckedValue);

            // Define output file path
            string outputPath = "CheckboxOnChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}