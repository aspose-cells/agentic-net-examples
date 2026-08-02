// Title: Export Workbook with DataBar Conditional Formatting to HTML and Verify Bar Widths – Aspose.Cells for .NET
// Description: Creates a workbook, fills column A with numeric values, applies a green DataBar conditional format to A1:A5, saves the file as HTML, extracts the rendered bar widths with a regular expression, computes expected percentages from the source data, and reports any deviation beyond a 0.5 % tolerance.
// Keywords: Aspose.Cells | C# | .NET | DataBar | conditional formatting | HTML export | verify bar width | width percentage regex | HtmlSaveOptions | CI regression test | visual validation
// Common Searches: Aspose.Cells export DataBar to HTML | verify DataBar width in generated HTML | C# extract CSS width from Aspose.Cells HTML output | conditional formatting bar length validation | automated test for Aspose.Cells HTML rendering
// Developer Intent: Generate an HTML report that includes DataBar conditional formatting and programmatically confirm that each bar’s rendered length accurately reflects its cell value.
// Use Cases: Produce numeric dashboards with DataBars in HTML and automatically validate visual fidelity before publishing. | Add a regression test to a CI pipeline that checks DataBar widths in the HTML output against source data. | Create a data‑driven email template where DataBar lengths must be verified to meet compliance standards.
// AI Prompts: Write C# code that adds a green DataBar conditional format to a range and saves the workbook as HTML using Aspose.Cells. | Provide a method to read the saved HTML, extract DataBar width percentages with a regex, and compare them to expected values derived from the worksheet. | Explain how to configure HtmlSaveOptions to render DataBars as solid fills instead of gradients for simpler width extraction.

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

// Creates a workbook, fills column A with numeric values, applies a green DataBar conditional format to A1:A5, saves the file as HTML, extracts the rendered bar widths with a regular expression, computes expected percentages from the source data, and reports any deviation beyond a 0.5 % tolerance.
class DataBarHtmlVerification
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample numeric data in column A (A1:A5)
            double[] values = { 10, 30, 50, 70, 90 };
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]);
            }

            // 3. Add a DataBar conditional formatting rule for the range A1:A5
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcs = sheet.ConditionalFormattings[cfIndex];

            // Define the area to which the formatting will be applied
            CellArea area = new CellArea
            {
                StartRow = 0,
                EndRow = values.Length - 1,
                StartColumn = 0,
                EndColumn = 0
            };
            fcs.AddArea(area);

            // Add the DataBar condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = fcs[conditionIdx];

            // Configure the DataBar (use automatic min/max and a visible color)
            DataBar dataBar = condition.DataBar;
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
            dataBar.Color = Color.Green;
            dataBar.ShowValue = true; // show the numeric value next to the bar

            // 4. Save the workbook as HTML (default DataBar render mode is Gradient)
            string htmlPath = "DataBarOutput.html";
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            workbook.Save(htmlPath, htmlOptions);

            // 5. Verify that the rendered bar lengths in the HTML correspond to the cell values
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file was not created: {htmlPath}");
                return;
            }

            string htmlContent = File.ReadAllText(htmlPath);

            // Extract all width percentages that belong to data bars.
            // The pattern looks for "width:" followed by a number (integer or decimal) and a percent sign.
            Regex widthRegex = new Regex(@"width\s*:\s*([0-9]*\.?[0-9]+)%", RegexOptions.IgnoreCase);
            MatchCollection matches = widthRegex.Matches(htmlContent);

            if (matches.Count != values.Length)
            {
                Console.WriteLine($"Expected {values.Length} data bar widths, but found {matches.Count}.");
                return;
            }

            // Compute expected percentages based on the min and max of the data set.
            double min = values[0];
            double max = values[0];
            foreach (double v in values)
            {
                if (v < min) min = v;
                if (v > max) max = v;
            }

            const double tolerance = 0.5; // allow half‑percent deviation due to rounding

            for (int i = 0; i < values.Length; i++)
            {
                double expected = (values[i] - min) / (max - min) * 100.0;
                double actual = double.Parse(matches[i].Groups[1].Value);
                if (Math.Abs(expected - actual) > tolerance)
                {
                    Console.WriteLine($"Row {i + 1}: Expected width {expected:F2}%, but found {actual:F2}%.");
                }
                else
                {
                    Console.WriteLine($"Row {i + 1}: Width verification passed ({actual:F2}% ≈ {expected:F2}%).");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
