// Title: Export Workbook with DataBar Conditional Formatting to HTML and Extract CSS Classes – Aspose.Cells for .NET
// Description: Creates a workbook, fills column A with values 1‑10, applies a green DataBar conditional format to A1:A10, saves the sheet as HTML with a separate .css file (ExportWorksheetCSSSeparately = true, ExcludeUnusedStyles = false), then reads the generated stylesheet and lists all CSS class selectors using a regular expression.
// Keywords: Aspose.Cells | C# | .NET | HTML export | separate CSS file | DataBar | conditional formatting | CSS class extraction | regex | ExportWorksheetCSSSeparately | ExcludeUnusedStyles
// Common Searches: Aspose.Cells export workbook to HTML with CSS file | DataBar conditional formatting HTML export .NET | how to extract CSS class names from Aspose.Cells output | save worksheet as HTML and keep conditional formatting styles | C# regex to list CSS selectors from generated stylesheet
// Developer Intent: Generate an HTML version of a workbook that retains DataBar conditional formatting, output the styles to a separate CSS file, and programmatically retrieve the CSS class names defined in that file.
// Use Cases: Build web reports that display Excel‑style DataBars and need the exact CSS classes for further UI customization. | Automate validation of conditional formatting by exporting to HTML and confirming the presence of the expected DataBar CSS selector. | Integrate Aspose‑generated styles into a custom front‑end framework by mapping extracted class names to existing design tokens.
// AI Prompts: Write C# code using Aspose.Cells to add a green DataBar conditional format, save the workbook as HTML with a separate CSS file, and extract all CSS class selectors from the generated stylesheet. | Explain which HtmlSaveOptions settings must be enabled so that the DataBar style is included in the exported CSS file. | Provide a robust C# regular expression that captures CSS class names (including hyphens and underscores) from the Aspose.Cells generated CSS content.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsDataBarCssExport
{
    // Creates a workbook, fills column A with values 1‑10, applies a green DataBar conditional format to A1:A10, saves the sheet as HTML with a separate .css file (ExportWorksheetCSSSeparately = true, ExcludeUnusedStyles = false), then reads the generated stylesheet and lists all CSS class selectors using a regular expression.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();                     // create
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            for (int i = 0; i < 10; i++)
            {
                // Values 1..10 in column A
                sheet.Cells[i, 0].PutValue(i + 1);
            }

            // ---------- Add DataBar conditional formatting ----------
            // Add an empty conditional formatting collection
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range A1:A10
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = 9,
                StartColumn = 0,
                EndColumn = 0
            };
            fcc.AddArea(area);

            // Add a DataBar condition
            int conditionIdx = fcc.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = fcc[conditionIdx];

            // Configure the DataBar (green bar, automatic min/max, hide values)
            DataBar dataBar = condition.DataBar;
            dataBar.Color = System.Drawing.Color.Green;
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
            dataBar.ShowValue = false;

            // ---------- Prepare HTML save options ----------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Export worksheet CSS into a separate .css file
            saveOptions.ExportWorksheetCSSSeparately = true;
            // Keep all styles (do not exclude unused ones) so the CSS file contains the DataBar class
            saveOptions.ExcludeUnusedStyles = false;

            // ---------- Define output paths ----------
            string outputFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "DataBarExport");

            Directory.CreateDirectory(outputFolder);

            string htmlPath = Path.Combine(outputFolder, "DataBarWorkbook.html");

            // ---------- Save workbook as HTML ----------
            workbook.Save(htmlPath, saveOptions);                  // save

            // ---------- Locate the generated CSS file ----------
            // When ExportWorksheetCSSSeparately = true, Aspose creates a CSS file named "sheet0.css"
            // (or similar). We'll search for *.css files in the output folder.
            string[] cssFiles = Directory.GetFiles(outputFolder, "*.css");
            if (cssFiles.Length == 0)
            {
                Console.WriteLine("No CSS file was generated.");
                return;
            }

            // For demonstration, read the first CSS file found.
            string cssFilePath = cssFiles[0];
            string cssContent = File.ReadAllText(cssFilePath);

            // ---------- Extract CSS class selectors ----------
            // Simple regex to capture class selectors like ".className {"
            Regex classRegex = new Regex(@"\.(\w[\w\-]*)\s*\{", RegexOptions.Compiled);
            MatchCollection matches = classRegex.Matches(cssContent);

            Console.WriteLine($"CSS file: {Path.GetFileName(cssFilePath)}");
            Console.WriteLine("Extracted CSS classes:");

            foreach (Match match in matches)
            {
                // Group 1 contains the class name without the leading dot
                Console.WriteLine("- " + match.Groups[1].Value);
            }

            // Optional: display the full CSS content for verification
            // Console.WriteLine("\nFull CSS content:\n" + cssContent);
        }
    }
}
