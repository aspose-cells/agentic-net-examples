// Title: Check Aspose.Cells HTML export for table class or id attributes with C# regex validation
// AI Prompts: Write C# code that saves a workbook to HTML using Aspose.Cells, reads the resulting file, and uses a regular expression to ensure every <table> element has either a class or an id attribute. | Update the HtmlSaveOptions to assign a custom TableCssIdPrefix (when the API supports it) and adjust the validation routine to confirm the prefix appears in each table's identifier. | Add comprehensive logging that enumerates any <table> tags lacking both class and id attributes and returns a pass/fail status for the HTML validation.
// Common Searches: aspocells htmlsaveoptions verify table id attribute c# | c# regex find table tags without class in exported html | how to set TableCssIdPrefix in Aspose.Cells HTML export | validate generated html tables contain identifiers using Aspose.Cells | check Aspose.Cells workbook to html for missing table identifiers
// Tags: Aspose.Cells HTML export table identifier check | C# regex verification of table tags | HtmlSaveOptions TableCssIdPrefix configuration | validate workbook to html table attributes | log missing table identifiers in exported html

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program creates a workbook, populates it with sample data, saves it as HTML via Aspose.Cells, reads the generated file, and uses a regular expression to confirm that each <table> tag includes either a class or an id attribute, reporting the overall validation result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Value1");
            sheet.Cells["B2"].PutValue("Value2");

            // Define a prefix that could be used for table IDs (if supported)
            string tableCssPrefix = "myPrefix";

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true
                // Note: TableCssIdPrefix is not available in the current Aspose.Cells version.
                // If a newer version is used, you can uncomment the line below:
                // TableCssIdPrefix = tableCssPrefix
            };

            // Export worksheet to HTML
            string htmlPath = "output.html";
            workbook.Save(htmlPath, htmlOptions);

            // Validate HTML content
            bool allValid = true;
            if (File.Exists(htmlPath))
            {
                string htmlContent;
                try
                {
                    htmlContent = File.ReadAllText(htmlPath);
                }
                catch (Exception readEx)
                {
                    Console.WriteLine($"Failed to read HTML file: {readEx.Message}");
                    allValid = false;
                    htmlContent = string.Empty;
                }

                // Find all <table> tags
                Regex tableRegex = new Regex(@"<table\b[^>]*>", RegexOptions.IgnoreCase);
                MatchCollection matches = tableRegex.Matches(htmlContent);

                if (matches.Count == 0)
                {
                    Console.WriteLine("No <table> tags found in the HTML output.");
                    allValid = false;
                }
                else
                {
                    foreach (Match match in matches)
                    {
                        string tag = match.Value;

                        // Prefer class attribute; fallback to id
                        string identifier = GetAttributeValue(tag, "class") ?? GetAttributeValue(tag, "id");

                        // If a specific prefix is required and the property is supported, validate it.
                        // Since TableCssIdPrefix is unavailable, we only ensure an identifier exists.
                        if (string.IsNullOrEmpty(identifier))
                        {
                            allValid = false;
                            Console.WriteLine("Table tag missing both class and id attributes.");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine($"HTML file not found: {htmlPath}");
                allValid = false;
            }

            // Output validation result
            Console.WriteLine(allValid
                ? "HTML validation succeeded."
                : "HTML validation failed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper to extract attribute value from a tag string
    private static string GetAttributeValue(string tag, string attributeName)
    {
        string pattern = $@"{attributeName}\s*=\s*[""']([^""']*)[""']";
        Match match = Regex.Match(tag, pattern, RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : null;
    }
}
