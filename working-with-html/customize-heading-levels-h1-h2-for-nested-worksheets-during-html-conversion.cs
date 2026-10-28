// Title: Adjust heading levels (keep first H1, convert others to H2) for multiple worksheets when exporting Excel to a single HTML file using Aspose.Cells for .NET
// Common Searches: How to keep only the first <h1> and turn the rest into <h2> after converting an Excel workbook to HTML with Aspose.Cells in C# | C# code to change extra heading tags when exporting multiple Excel sheets to one HTML file using Aspose.Cells | Setting HtmlSaveOptions to export all worksheets and then adjusting heading levels in the resulting HTML | Using reflection to apply UseCss formatting option in Aspose.Cells HTML export for .NET | Regex pattern to replace subsequent h1 elements with h2 in Aspose.Cells generated HTML
// Tags: Aspose.Cells export all worksheets to single HTML | modify heading tags in Aspose.Cells HTML output | C# regex replace h1 with h2 in HTML | HtmlSaveOptions ExportActiveWorksheetOnly false | use reflection to enable HtmlFormattingOptions UseCss

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, uses HtmlSaveOptions with ExportActiveWorksheetOnly set to false to export all worksheets into one HTML file, optionally enables CSS formatting via reflection, reads the generated HTML from a memory stream, and applies a regular expression that preserves the first <h1> tag while converting subsequent <h1> tags to <h2>. The final HTML is written to disk after ensuring the output directory exists.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Configure HTML save options (export all worksheets)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false
            };

            // Attempt to enable CSS formatting if the property exists in this version
            try
            {
                var prop = typeof(HtmlSaveOptions).GetProperty("HtmlFormattingOptions");
                if (prop != null && prop.CanWrite)
                {
                    var enumType = prop.PropertyType;
                    var useCss = Enum.Parse(enumType, "UseCss");
                    prop.SetValue(htmlOptions, useCss);
                }
            }
            catch
            {
                // Property not supported – continue with defaults
            }

            // Save workbook to a memory stream as HTML
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;

                // Read generated HTML
                string htmlContent = new StreamReader(ms).ReadToEnd();

                // Adjust heading levels: keep first <h1>, change others to <h2>
                int headingCount = 0;
                htmlContent = Regex.Replace(
                    htmlContent,
                    "<h1>(.*?)</h1>",
                    match =>
                    {
                        headingCount++;
                        return headingCount == 1
                            ? match.Value
                            : $"<h2>{match.Groups[1].Value}</h2>";
                    },
                    RegexOptions.Singleline);

                // Ensure the output directory exists
                try
                {
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create output directory: {ex.Message}");
                }

                // Write final HTML to disk
                File.WriteAllText(outputPath, htmlContent);
                Console.WriteLine($"HTML saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
