// Title: Export an Excel workbook to HTML with Aspose.Cells while controlling hidden rows using HtmlSaveOptions.ExportHiddenWorksheet
// AI Prompts: Write a C# program that creates a workbook, hides specific rows, and saves the workbook to HTML twice—once with default HtmlSaveOptions and once with ExportHiddenWorksheet set to true. | Add code that reads the two generated HTML files, checks for the hidden‑row values, and prints verification results indicating whether hidden rows were included.
// Common Searches: c# aspocells export hidden rows to html example | how to use HtmlSaveOptions.ExportHiddenWorksheet in Aspose.Cells | exclude hidden rows when saving workbook as html with Aspose.Cells | verify hidden rows are omitted in html output using Aspose.Cells C# | Aspose.Cells default behavior for hidden rows in html export
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# export workbook to HTML hidden rows | HTML export hidden rows Aspose.Cells | verify hidden row presence in exported HTML | SaveFormat.Html example Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, hides rows 2 and 4, then saves the workbook to HTML twice—first with default options (hidden rows omitted) and second with HtmlSaveOptions.ExportHiddenWorksheet enabled (hidden rows included). It reads both HTML files and programmatically verifies whether the hidden‑row values appear, outputting the verification results.
class ExportHiddenRowsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data in column A (rows 1 to 5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue($"Row {i + 1}");
            }

            // Hide rows 2 and 4 (zero‑based index: rows 1 and 3)
            sheet.Cells.Rows[1].IsHidden = true; // Row 2
            sheet.Cells.Rows[3].IsHidden = true; // Row 4

            // -----------------------------------------------------------------
            // Export to HTML with default options (hidden rows are omitted)
            // -----------------------------------------------------------------
            string htmlPathDefault = "Workbook_Default.html";
            workbook.Save(htmlPathDefault, SaveFormat.Html);

            // -----------------------------------------------------------------
            // Export to HTML with ExportHiddenWorksheet = true (hidden rows are included)
            // -----------------------------------------------------------------
            string htmlPathIncludeHidden = "Workbook_IncludeHidden.html";
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportHiddenWorksheet = true // Include hidden rows/columns
            };
            workbook.Save(htmlPathIncludeHidden, htmlOptions);

            // -----------------------------------------------------------------
            // Verify the results by checking the presence of hidden row values
            // -----------------------------------------------------------------
            string htmlDefault = File.Exists(htmlPathDefault) ? File.ReadAllText(htmlPathDefault) : string.Empty;
            string htmlIncludeHidden = File.Exists(htmlPathIncludeHidden) ? File.ReadAllText(htmlPathIncludeHidden) : string.Empty;

            // Values that belong to hidden rows
            const string hiddenRow2Value = "Row 2";
            const string hiddenRow4Value = "Row 4";

            // Check default export (should NOT contain hidden rows)
            bool defaultContainsHiddenRow2 = htmlDefault.Contains(hiddenRow2Value);
            bool defaultContainsHiddenRow4 = htmlDefault.Contains(hiddenRow4Value);

            // Check export with hidden rows included (should contain hidden rows)
            bool includeHiddenContainsRow2 = htmlIncludeHidden.Contains(hiddenRow2Value);
            bool includeHiddenContainsRow4 = htmlIncludeHidden.Contains(hiddenRow4Value);

            Console.WriteLine("Verification Results:");
            Console.WriteLine($"Default export contains hidden Row 2: {defaultContainsHiddenRow2}");
            Console.WriteLine($"Default export contains hidden Row 4: {defaultContainsHiddenRow4}");
            Console.WriteLine($"Export with ExportHiddenWorksheet=true contains hidden Row 2: {includeHiddenContainsRow2}");
            Console.WriteLine($"Export with ExportHiddenWorksheet=true contains hidden Row 4: {includeHiddenContainsRow4}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
