using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsCheckboxOnChartSheet
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a worksheet (used as a chart sheet substitute) to the workbook
                Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet1");
                int sheetIdx = chartSheet.Index; // index of the newly added sheet (optional)

                // Add a checkbox to the sheet
                // Parameters: upperLeftRow, upperLeftColumn, height (pixels), width (pixels)
                int checkBoxIdx = chartSheet.CheckBoxes.Add(2, 2, 20, 100);
                CheckBox checkBox = chartSheet.CheckBoxes[checkBoxIdx];

                // Set checkbox properties: label and default state (unchecked)
                checkBox.Text = "Select";
                checkBox.Value = false; // unchecked

                // Define output file path
                string outputPath = "ChartSheetWithCheckBox.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}