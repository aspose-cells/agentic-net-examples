// Title: Import an HTML table with colspan and rowspan into an Excel worksheet as merged cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load HTML file and keep colspan rowspan merged cells | How to preserve merged cells when converting HTML tables to Excel with Aspose.Cells | Aspose.Cells HtmlLoadOptions example for merging cells from an HTML table | Convert HTML table with merged cells to XLSX using Aspose.Cells .NET | List merged ranges after importing HTML into a workbook with Aspose.Cells
// Tags: Aspose.Cells HtmlLoadOptions import HTML table | convert HTML colspan rowspan to Excel merged cells | enumerate merged ranges in workbook C# | save workbook as XLSX using Aspose.Cells | console argument validation for file paths C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The console program validates input arguments, checks the source HTML file, creates the output directory if needed, and loads the HTML using Aspose.Cells' HtmlLoadOptions, which automatically converts <table> elements with colspan and rowspan into merged cells. It then iterates through each worksheet to output detected merged ranges and finally saves the workbook as an XLSX file, handling any errors that occur.
class Program
{
    static void Main(string[] args)
    {
        // Validate arguments
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: Program <inputHtmlPath> <outputExcelPath>");
            return;
        }

        string htmlPath = args[0];
        string excelPath = args[1];

        // Verify input HTML file exists
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
            return;
        }

        // Ensure the output directory exists
        try
        {
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
        }
        catch (Exception dirEx)
        {
            Console.WriteLine($"Error creating output directory: {dirEx.Message}");
            return;
        }

        try
        {
            // Load the HTML file. Aspose.Cells automatically converts <table> elements,
            // respecting colspan and rowspan attributes, into merged cells.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Optional: iterate through all worksheets and list merged ranges for verification.
            Console.WriteLine("Merged cells detected after HTML import:");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");
                bool anyMerged = false;
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsMerged)
                    {
                        // The top‑left cell of a merged region reports IsMerged = true.
                        // Use cell.GetMergedRange() to obtain the full range.
                        AsposeRange mergedRange = cell.GetMergedRange();
                        Console.WriteLine($"  {mergedRange.RefersTo} (Rows: {mergedRange.RowCount}, Columns: {mergedRange.ColumnCount})");
                        anyMerged = true;
                    }
                }
                if (!anyMerged)
                {
                    Console.WriteLine("  (No merged cells in this worksheet)");
                }
            }

            // Save the workbook in XLSX format.
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Successfully saved Excel file to '{excelPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the HTML file:");
            Console.WriteLine(ex.Message);
        }
    }
}
