using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class DataBarHtmlExportAndVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column A (values 10, 50, 100)
            double[] values = { 10, 50, 100 };
            for (int i = 0; i < values.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(values[i]);
            }

            // Add a DataBar conditional formatting to the range A1:A3
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

            // Add a DataBar condition
            int conditionIdx = fcs.AddCondition(FormatConditionType.DataBar);
            FormatCondition condition = fcs[conditionIdx];

            // Configure the DataBar (automatic min/max, green color, show values)
            DataBar dataBar = condition.DataBar;
            dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin;
            dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax;
            dataBar.Color = Color.Green;
            dataBar.ShowValue = true; // show numeric value beside the bar

            // Save the workbook as HTML with DataBar rendered as a bar (default rendering)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // The Bar rendering mode may not be available in older versions; using default.
            // If needed, uncomment the line below and ensure the enum contains 'Bar'.
            // htmlOptions.DataBarRenderMode = DataBarRenderMode.Bar;
            string htmlPath = "DataBarExport.html";
            workbook.Save(htmlPath, htmlOptions);

            // ---------------- Verification ----------------
            // Ensure the HTML file was created before reading
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Verification failed: '{htmlPath}' was not created.");
                return;
            }

            // Load the generated HTML
            string htmlContent = File.ReadAllText(htmlPath);

            // Determine the maximum value (used by the automatic DataBar scaling)
            double maxValue = double.MinValue;
            foreach (double v in values)
                if (v > maxValue) maxValue = v;

            // Regex to capture the width percentage of each data bar cell.
            // Aspose.Cells renders a data bar as a <div> with style "width:XX%".
            Regex widthRegex = new Regex(@"<div[^>]*style=[""'][^""']*width\s*:\s*(\d+)%", RegexOptions.IgnoreCase);

            // Find all matches in the HTML (order corresponds to cell order)
            MatchCollection matches = widthRegex.Matches(htmlContent);
            if (matches.Count != values.Length)
            {
                Console.WriteLine("Verification failed: number of data bar elements does not match cell count.");
                return;
            }

            // Compare each bar length with the expected percentage
            bool allMatch = true;
            for (int i = 0; i < values.Length; i++)
            {
                // Expected percentage (rounded to nearest integer)
                int expectedPercent = (int)Math.Round(values[i] / maxValue * 100);
                int actualPercent = int.Parse(matches[i].Groups[1].Value);

                Console.WriteLine($"Row {i + 1}: Value={values[i]}, Expected%={expectedPercent}, Actual%={actualPercent}");

                if (expectedPercent != actualPercent)
                    allMatch = false;
            }

            Console.WriteLine(allMatch
                ? "All data bar lengths correctly reflect cell values."
                : "Some data bar lengths do not match the expected values.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}