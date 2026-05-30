using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsDataBarHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (rows 1-7)
            for (int i = 0; i < 7; i++)
            {
                sheet.Cells[i, 0].PutValue((i + 1) * 10); // 10,20,...,70
            }

            // Add an empty conditional formatting collection to the worksheet
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range for the DataBar (A1:A7)
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 6,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a DataBar condition
            int conditionIdx = fcc.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = fcc[conditionIdx];

            // Configure the DataBar appearance
            DataBar dataBar = condition.DataBar;
            dataBar.Color = System.Drawing.Color.Orange;               // Bar color
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum value
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum value
            dataBar.ShowValue = true;                                   // Show cell values
            dataBar.BarBorder.Type = DataBarBorderType.Solid;            // Solid border
            dataBar.BarBorder.Color = System.Drawing.Color.DarkBlue;    // Border color
            dataBar.AxisPosition = DataBarAxisPosition.Midpoint;       // Axis position

            // Prepare HTML save options to export worksheet CSS separately
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportWorksheetCSSSeparately = true; // Generates a separate .css file
            saveOptions.DisableCss = false;                 // Keep CSS (not inline only)

            // Define output directory and file names
            string outputDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "AsposeDataBarHtml");
            Directory.CreateDirectory(outputDir);

            string htmlPath = Path.Combine(outputDir, "DataBarOutput.html");
            // Save the workbook as HTML
            workbook.Save(htmlPath, saveOptions);

            // After saving, locate the generated CSS file (usually named sheet0.css)
            string[] cssFiles = Directory.GetFiles(outputDir, "*.css");
            if (cssFiles.Length > 0)
            {
                // Read and display the CSS content for analysis
                string cssContent = File.ReadAllText(cssFiles[0]);
                Console.WriteLine("=== Generated CSS ===");
                Console.WriteLine(cssContent);
            }
            else
            {
                Console.WriteLine("No CSS file was generated.");
            }

            // Clean up
            workbook.Dispose();
        }
    }
}