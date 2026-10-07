// Title: Validate that Aspose.Cells HTML export uses a custom TableId prefix for all <table> elements in C#
// AI Prompts: Generate C# code that creates a workbook, applies a specific TableId prefix via HtmlSaveOptions (using reflection for version compatibility), saves it as HTML, and then uses a case‑insensitive regular expression to confirm each <table> tag's id starts with that prefix. | Write a C# example that reads the HTML file produced by Aspose.Cells, extracts all table id attributes, and reports any identifiers that do not begin with the defined prefix.
// Common Searches: aspocells html export modify table identifiers c# | c# regex verify table id attributes in aspocells generated html | use reflection to set HtmlSaveOptions properties for Aspose.Cells versions | check aspocells exported html for correct table id naming convention | c# programmatically validate html tables after Aspose.Cells save
// Tags: Aspose.Cells HtmlSaveOptions TableIdPrefix | C# regex HTML table ID validation | Aspose.Cells HTML export custom table identifiers | reflection configure HtmlExportOptions C# | verify generated HTML table attributes Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlValidation
{
    // The sample creates an Aspose.Cells workbook, sets a custom TableId prefix on HTML tables via reflection‑compatible HtmlSaveOptions, saves the workbook as HTML, then reads the output file and uses a case‑insensitive regex to ensure every <table> element's id attribute begins with the expected prefix, reporting any mismatches.
    class Program
    {
        static void Main(string[] args)
        {
            // Define output HTML path
            const string htmlPath = "Sample.html";

            // Create a new workbook, populate data and save as HTML
            try
            {
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Row1Col1");
                sheet.Cells["B2"].PutValue("Row1Col2");

                // Configure HTML save options with a TableId prefix
                var saveOptions = new HtmlSaveOptions(SaveFormat.Html);
                const string tableIdPrefix = "tbl_";

                // Use reflection to set TableIdPrefix (covers different Aspose.Cells versions)
                try
                {
                    var htmlExportProp = typeof(HtmlSaveOptions).GetProperty("HtmlExportOptions");
                    if (htmlExportProp != null)
                    {
                        var htmlExport = htmlExportProp.GetValue(saveOptions);
                        var prefixProp = htmlExport?.GetType().GetProperty("TableIdPrefix");
                        if (prefixProp != null && prefixProp.CanWrite)
                        {
                            prefixProp.SetValue(htmlExport, tableIdPrefix);
                        }
                    }
                }
                catch
                {
                    // If reflection fails, continue without custom prefix
                }

                // Save the workbook as HTML
                workbook.Save(htmlPath, saveOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during workbook creation or HTML saving: {ex.Message}");
                return;
            }

            // Validate the generated HTML
            try
            {
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"HTML file not found: {htmlPath}");
                    return;
                }

                string htmlContent = File.ReadAllText(htmlPath);

                // Regex to match <table ... id="..."> (case‑insensitive)
                var tableRegex = new Regex(@"<table[^>]*\sid\s*=\s*[""']([^""']+)[""'][^>]*>", RegexOptions.IgnoreCase);
                var matches = tableRegex.Matches(htmlContent);

                bool allTablesValid = true;
                const string expectedPrefix = "tbl_";

                foreach (Match match in matches)
                {
                    string idValue = match.Groups[1].Value;
                    if (!idValue.StartsWith(expectedPrefix, StringComparison.Ordinal))
                    {
                        allTablesValid = false;
                        Console.WriteLine($"Invalid table ID: {idValue}");
                    }
                }

                Console.WriteLine(allTablesValid
                    ? "All table elements have correct TableId prefixes."
                    : "One or more table elements have incorrect TableId prefixes.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during HTML validation: {ex.Message}");
            }
        }
    }
}
