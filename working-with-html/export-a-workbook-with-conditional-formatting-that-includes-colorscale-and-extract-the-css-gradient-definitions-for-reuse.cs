using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using System.Drawing;

class ColorScaleCssExtractor
{
    static void Main()
    {
        // 1. Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // 2. Populate the worksheet with sample numeric data (10x10 matrix)
        for (int i = 0; i < 10; i++)
        {
            for (int j = 0; j < 10; j++)
            {
                sheet.Cells[i, j].PutValue(i * j);
            }
        }

        // 3. Add a ColorScale conditional formatting rule
        int cfIndex = sheet.ConditionalFormattings.Add();                     // create a new ConditionalFormatting collection
        FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

        // Define the range A1:J10 for the conditional formatting
        CellArea area = new CellArea { StartRow = 0, EndRow = 9, StartColumn = 0, EndColumn = 9 };
        fcc.AddArea(area);

        // Add a ColorScale condition (type = ColorScale)
        int conditionIdx = fcc.AddCondition(FormatConditionType.ColorScale);
        FormatCondition fc = fcc[conditionIdx];

        // Configure a 3‑color scale: Red (min) → Yellow (mid) → Green (max)
        fc.ColorScale.Is3ColorScale = true;

        fc.ColorScale.MinColor = Color.Red;
        fc.ColorScale.MinCfvo.Type = FormatConditionValueType.Min;          // minimum value

        fc.ColorScale.MidColor = Color.Yellow;
        fc.ColorScale.MidCfvo.Type = FormatConditionValueType.Percentile; // percentile value
        fc.ColorScale.MidCfvo.Value = 50;                                   // 50th percentile

        fc.ColorScale.MaxColor = Color.Green;
        fc.ColorScale.MaxCfvo.Type = FormatConditionValueType.Max;          // maximum value

        // 4. Save the workbook as HTML with CSS exported separately
        string htmlPath = "ColorScale.html";
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true; // generate a separate CSS file per worksheet
        workbook.Save(htmlPath, saveOptions);

        // 5. Locate the generated CSS file (default name: sheet0.css in the same folder)
        string cssFolder = Path.GetDirectoryName(Path.GetFullPath(htmlPath));
        string cssPath = Path.Combine(cssFolder, "sheet0.css");

        if (File.Exists(cssPath))
        {
            // 6. Read the CSS content
            string cssContent = File.ReadAllText(cssPath);

            // 7. Extract gradient definitions (e.g., linear-gradient) using a regular expression
            Regex gradientRegex = new Regex(@"background\s*:\s*linear-gradient\([^\)]+\)", RegexOptions.IgnoreCase);
            MatchCollection matches = gradientRegex.Matches(cssContent);

            Console.WriteLine("Extracted CSS gradient definitions:");
            foreach (Match match in matches)
            {
                // Output each gradient rule (trimmed) followed by a semicolon for completeness
                Console.WriteLine(match.Value.Trim() + ";");
            }
        }
        else
        {
            Console.WriteLine("CSS file not found: " + cssPath);
        }
    }
}