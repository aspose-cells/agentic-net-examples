// Title: Export a workbook with DataBar conditional formatting to HTML and validate bar widths against cell values using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, applies a DataBar conditional format to a range, saves the workbook as HTML to a memory stream, and programmatically verifies that each bar's CSS width matches the corresponding cell's numeric value. | Write a C# routine with Aspose.Cells that exports a worksheet containing DataBar conditional formatting to HTML, extracts the width percentages of the rendered bars using regex, and compares them to the expected percentages calculated from the cell data.
// Common Searches: how to check data bar width percentages after exporting to HTML with Aspose.Cells .NET | C# verify that Aspose.Cells DataBar conditional formatting renders correct bar lengths in HTML | extract CSS width of DataBar from Aspose.Cells generated HTML for validation | compare numeric cell values to rendered data bar lengths in Aspose.Cells HTML output
// Tags: Aspose.Cells export workbook to HTML with DataBar | C# validate DataBar CSS width percentages | conditional formatting DataBar HTML rendering verification | parse generated HTML for data bar widths using regex | compare cell values to data bar lengths in Aspose.Cells

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook, fills column A with numeric values, adds a DataBar conditional formatting rule, saves the workbook as HTML to a memory stream, parses the resulting HTML to extract each data bar's width percentage, and checks that the percentages correspond to the underlying cell values within a small tolerance.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Populate sample numeric data in range A1:A5
            double[] values = { 10, 30, 50, 70, 90 };
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]); // Column A (index 0)
            }

            // 3. Apply DataBar conditional formatting to the range A1:A5
            int firstRow = 0;
            int firstColumn = 0;
            int totalRows = values.Length;
            int totalColumns = 1;

            // Ensure a ConditionalFormatting collection exists and add a new rule
            ConditionalFormattingCollection cfCollection = sheet.ConditionalFormattings;
            int cfIndex = cfCollection.Add(); // adds a new ConditionalFormatting object

            // Use dynamic to avoid compile‑time dependency on ConditionalFormatting type
            dynamic cf = cfCollection[cfIndex];

            // Define the area the rule applies to
            cf.AddArea(firstRow, firstColumn, totalRows, totalColumns);

            // Add a DataBar condition
            FormatCondition dataBarCondition = cf.AddCondition(FormatConditionType.DataBar);
            dataBarCondition.DataBar.ShowValue = true; // show the numeric value inside the bar

            // 4. Save the workbook as HTML to a memory stream
            using (MemoryStream htmlStream = new MemoryStream())
            {
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0;

                // 5. Load the generated HTML as a string for verification
                string html = Encoding.UTF8.GetString(htmlStream.ToArray());

                // 6. Extract width percentages of data bars using regex
                Regex widthRegex = new Regex(
                    @"<tr[^>]*>.*?<td[^>]*>.*?width\s*:\s*([\d\.]+)%", 
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                MatchCollection matches = widthRegex.Matches(html);

                // 7. Determine the min and max values used for the DataBar
                double minValue = values.Min();
                double maxValue = values.Max();

                bool allMatch = true;
                for (int i = 0; i < values.Length; i++)
                {
                    if (i >= matches.Count)
                    {
                        Console.WriteLine($"Missing data bar for row {i + 1}.");
                        allMatch = false;
                        continue;
                    }

                    // Parse the actual percentage from the HTML
                    string percentStr = matches[i].Groups[1].Value;
                    if (!double.TryParse(percentStr, out double actualPercent))
                    {
                        Console.WriteLine($"Failed to parse width percentage for row {i + 1}.");
                        allMatch = false;
                        continue;
                    }

                    // Compute expected percentage based on value proportion
                    double expectedPercent = (values[i] - minValue) / (maxValue - minValue) * 100.0;

                    // Allow a small tolerance due to rounding in HTML rendering
                    const double tolerance = 1.0; // 1%
                    if (Math.Abs(actualPercent - expectedPercent) > tolerance)
                    {
                        Console.WriteLine(
                            $"Row {i + 1}: Value={values[i]}, Expected≈{expectedPercent:F1}%, Actual={actualPercent}% -> MISMATCH");
                        allMatch = false;
                    }
                    else
                    {
                        Console.WriteLine(
                            $"Row {i + 1}: Value={values[i]}, Expected≈{expectedPercent:F1}%, Actual={actualPercent}% -> OK");
                    }
                }

                Console.WriteLine(allMatch
                    ? "All data bar lengths match cell values."
                    : "Some data bar lengths do not match.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
