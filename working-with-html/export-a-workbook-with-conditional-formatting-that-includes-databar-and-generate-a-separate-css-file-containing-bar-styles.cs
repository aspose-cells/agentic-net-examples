using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsDataBarHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample numeric data in column A (A1:A10)
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i * 10 + 5);
            }

            // Add a DataBar conditional formatting rule for the range A1:A10
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the cell area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add the DataBar condition
            int conditionIndex = fcs.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = fcs[conditionIndex];

            // Configure the DataBar appearance
            DataBar dataBar = condition.DataBar;
            dataBar.Color = Color.Orange;
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
            dataBar.ShowValue = true;
            dataBar.BarBorder.Type = DataBarBorderType.Solid;
            dataBar.BarBorder.Color = Color.DarkRed;
            dataBar.BarFillType = DataBarFillType.Solid;

            // Prepare HTML save options to export CSS separately
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportWorksheetCSSSeparately = true;
            // Optional: exclude unused styles to keep the CSS file small
            saveOptions.ExcludeUnusedStyles = true;

            // Determine output paths
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AsposeDataBarHtml");
            Directory.CreateDirectory(outputDir);
            string htmlPath = Path.Combine(outputDir, "DataBar.html");

            // Save the workbook as HTML; a separate .css file will be generated in the same folder
            workbook.Save(htmlPath, saveOptions);

            Console.WriteLine("HTML and separate CSS files have been saved to: " + outputDir);
        }
    }
}