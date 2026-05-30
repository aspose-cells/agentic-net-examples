using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsColorScaleHtmlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (0..9) in range A1:J10
            for (int row = 0; row <= 9; row++)
            {
                for (int col = 0; col <= 9; col++)
                {
                    sheet.Cells[row, col].PutValue(row * col);
                }
            }

            // Add a three‑color scale conditional formatting to the same range
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the area to which the formatting applies
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 9
            };
            fcs.AddArea(area);

            // Add the ColorScale condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.ColorScale);
            FormatCondition fc = fcs[conditionIdx];

            // Configure a 3‑color scale: Min = Red, Mid = Yellow (50th percentile), Max = Green
            fc.ColorScale.Is3ColorScale = true;
            fc.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;
            fc.ColorScale.MinColor = Color.Red;

            fc.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile;
            fc.ColorScale.MidCfvo.Value = 50;
            fc.ColorScale.MidColor = Color.Yellow;

            fc.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;
            fc.ColorScale.MaxColor = Color.Green;

            // Prepare HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export each worksheet's CSS separately for easier inspection (optional)
                ExportWorksheetCSSSeparately = true,
                // Keep all styles so we can see the generated colors
                ExcludeUnusedStyles = false
            };

            // Define output HTML path
            string htmlPath = Path.Combine(Environment.CurrentDirectory, "ColorScaleDemo.html");

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the generated HTML contains the expected color codes
            string htmlContent = File.ReadAllText(htmlPath);

            // Convert the .NET Color values to HTML hex strings (e.g., #FF0000)
            string redHex = ColorTranslator.ToHtml(Color.Red);
            string yellowHex = ColorTranslator.ToHtml(Color.Yellow);
            string greenHex = ColorTranslator.ToHtml(Color.Green);

            bool containsRed = htmlContent.IndexOf(redHex, StringComparison.OrdinalIgnoreCase) >= 0;
            bool containsYellow = htmlContent.IndexOf(yellowHex, StringComparison.OrdinalIgnoreCase) >= 0;
            bool containsGreen = htmlContent.IndexOf(greenHex, StringComparison.OrdinalIgnoreCase) >= 0;

            Console.WriteLine("HTML file saved to: " + htmlPath);
            Console.WriteLine("Verification of gradient colors in HTML:");
            Console.WriteLine($" - Contains {redHex}: {containsRed}");
            Console.WriteLine($" - Contains {yellowHex}: {containsYellow}");
            Console.WriteLine($" - Contains {greenHex}: {containsGreen}");

            // Simple result output
            if (containsRed && containsYellow && containsGreen)
                Console.WriteLine("Gradient colors are present in the HTML output.");
            else
                Console.WriteLine("One or more gradient colors are missing in the HTML output.");
        }
    }
}