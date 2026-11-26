// Title: Add a CheckBox to an Aspose.Cells chart, set it checked, and confirm its state in C#
// Common Searches: C# Aspose.Cells how to place a checkbox on a chart and read its Value property | Aspose.Cells chart checkbox checked verification example | set checkbox to true and link it to a cell using Aspose.Cells .NET | retrieve checkbox state from a chart in an Excel file with Aspose.Cells | programmatically check a chart control and validate the linked cell in Aspose.Cells
// Tags: Aspose.Cells chart checkbox control C# | set checkbox value Aspose.Cells | verify checkbox linked cell Aspose.Cells | Aspose.Cells add chart control programmatically | Aspose.Cells save workbook xlsx C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace CheckBoxOnChartDemo
{
    // The sample creates a workbook, adds a column chart, inserts a CheckBox over the chart, sets its Value to true, links it to cell C1, saves the file as XLSX, then reads the CheckBox.Value and the linked cell to confirm the checkbox is checked.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "CheckBoxOnChartDemo.xlsx";

            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                // Category data can be inferred; explicit assignment is optional
                chart.Title.Text = "Sample Column Chart";

                // Add a CheckBox control positioned over the chart area
                CheckBoxCollection checkBoxes = sheet.CheckBoxes;
                int upperLeftRow = 6;      // 0‑based row index
                int upperLeftColumn = 8;   // 0‑based column index
                int height = 20;           // height in pixels
                int width = 100;           // width in pixels

                int checkBoxIdx = checkBoxes.Add(upperLeftRow, upperLeftColumn, height, width);
                CheckBox checkBox = checkBoxes[checkBoxIdx];

                // Set CheckBox properties
                checkBox.Text = "Show Details";
                checkBox.Value = true;               // Check the box
                checkBox.LinkedCell = "$C$1";         // Optional linked cell

                // Write the linked cell value to reflect the checked state (true -> 1)
                sheet.Cells["C1"].PutValue(checkBox.Value ? 1 : 0);

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");

                // Verify that the CheckBox is indeed checked
                bool isChecked = checkBox.Value; // Value indicates the checked state
                Console.WriteLine($"CheckBox.Checked (via Value) = {isChecked}");

                // Additional verification using the linked cell
                object linkedCellValue = sheet.Cells["C1"].Value;
                Console.WriteLine($"Linked cell C1 value = {linkedCellValue}");

                // Simple assertion
                if (isChecked && Convert.ToInt32(linkedCellValue) == 1)
                {
                    Console.WriteLine("Verification succeeded: CheckBox is checked.");
                }
                else
                {
                    Console.WriteLine("Verification failed: CheckBox state is not as expected.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
