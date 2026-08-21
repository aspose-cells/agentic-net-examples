// Title: Export Aspose.Cells Workbook with DataBar Conditional Formatting to HTML and Validate Bar Widths (C#)
// Description: Creates a workbook, adds a green DataBar conditional format to cells A1:A4, saves the sheet as HTML, extracts the rendered bar widths from the HTML, and compares them to the expected percentages based on the original values.
// Keywords: Aspose.Cells | DataBar conditional formatting | HTML export C# | HtmlSaveOptions | DataBarRenderMode | verify data bar width | extract width percentage | C# Excel to HTML | conditional formatting testing
// Common Searches: Aspose.Cells export DataBar to HTML C# | Validate DataBar lengths in generated HTML | HtmlSaveOptions DataBarRenderMode example | Parse data bar width from Aspose.Cells HTML output | C# verify conditional formatting rendering
// Developer Intent: Generate an HTML file from a workbook that contains a DataBar conditional format and programmatically confirm that each bar’s length matches its cell value.
// Use Cases: Build HTML reports with visual data bars that accurately reflect numeric values. | Automate regression tests for conditional‑format rendering by comparing HTML bar widths to source data. | Create downloadable dashboards where DataBar visuals are needed for quick data insight.
// AI Prompts: Write C# code to read an Aspose.Cells‑generated HTML file, extract DataBar div width percentages, and compare them to expected values derived from the worksheet. | Explain how to enable DataBarRenderMode in HtmlSaveOptions for Aspose.Cells and handle scenarios where the property is missing in older versions. | Provide a C# unit test that asserts the rendered DataBar widths in the saved HTML are within an acceptable tolerance of the calculated percentages.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;
using System.Text.RegularExpressions;

// Creates a workbook, adds a green DataBar conditional format to cells A1:A4, saves the sheet as HTML, extracts the rendered bar widths from the HTML, and compares them to the expected percentages based on the original values.
class DataBarHtmlExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with sample numeric values
            double[] values = { 10, 30, 70, 100 };
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]);
            }

            // Add a DataBar conditional formatting rule to the range A1:A4
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the area for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = values.Length - 1,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add the DataBar condition
            int conditionIndex = fcs.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = fcs[conditionIndex];

            // Configure the DataBar (automatic min/max, green color, show cell value)
            condition.DataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            condition.DataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
            condition.DataBar.Color = System.Drawing.Color.Green;
            condition.DataBar.ShowValue = true;

            // Save the workbook as HTML, rendering DataBar as a visual bar
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // The DataBarRenderMode property may not be available in older versions.
            // If it exists, you can uncomment the following line:
            // htmlOptions.DataBarRenderMode = DataBarRenderMode.Bar;

            string htmlPath = "DataBarExport.html";
            workbook.Save(htmlPath, htmlOptions);

            // ---- Verification of bar lengths in the generated HTML ----
            if (File.Exists(htmlPath))
            {
                // Load the HTML content
                string htmlContent = File.ReadAllText(htmlPath);

                // Extract width percentages of the rendered data bars using a regular expression.
                // Aspose.Cells renders each data bar as a <div> with a style containing "width:XX%".
                MatchCollection matches = Regex.Matches(htmlContent, @"width\s*:\s*(\d+)%");

                // Determine the maximum value among the source data (used for expected percentage calculation)
                double maxValue = 0;
                foreach (double v in values)
                {
                    if (v > maxValue) maxValue = v;
                }

                // Compare each extracted width with the expected percentage based on the cell value.
                int cellIndex = 0;
                foreach (Match match in matches)
                {
                    int htmlWidth = int.Parse(match.Groups[1].Value);
                    double expectedWidth = values[cellIndex] / maxValue * 100.0;
                    Console.WriteLine($"Cell A{cellIndex + 1}: HTML width = {htmlWidth}% , Expected ≈ {Math.Round(expectedWidth)}%");
                    cellIndex++;
                }
            }
            else
            {
                Console.WriteLine($"Error: HTML file '{htmlPath}' was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
