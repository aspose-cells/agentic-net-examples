// Title: Remove tooltip attributes from short‑text cells after exporting Excel to HTML with AddTooltipText enabled using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, saves it to HTML with HtmlSaveOptions.AddTooltipText = true, then post‑processes the HTML to delete the title attribute from <td> elements whose displayed text length is below a specified threshold. | Write a C# regular expression that matches <td> tags containing a title attribute and returns the tag without the title when the inner cell text is shorter than five characters. | Adapt the HTML post‑processing step to preserve tooltip attributes for cells that contain line breaks or HTML markup while still removing them from plain short‑text cells.
// Common Searches: aspocells c# remove tooltip attribute from short cells in html export | how to filter AddTooltipText output for cells with less than 5 characters | regex to delete title attribute from generated html table cells aspocells | post processing html saved from excel to drop tooltips on short text cells
// Tags: Aspose.Cells HTML tooltip removal | C# regex strip title attribute from td | AddTooltipText short text filtering | post‑process generated HTML from Excel | Excel to HTML tooltip suppression .NET

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook, saves it as HTML with AddTooltipText turned on, then reads the resulting file and uses a regular expression to remove the title attribute from <td> elements whose inner text length is less than a defined threshold, writing the cleaned HTML back to disk.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";
            const int shortTextThreshold = 5;

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options with tooltip generation enabled
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                AddTooltipText = true // Enable tooltip attributes
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);

            // Post‑process the generated HTML to remove tooltips from short‑text cells
            if (File.Exists(outputPath))
            {
                string htmlContent = File.ReadAllText(outputPath);

                // Regex to find <td> elements with a title attribute
                string pattern = @"<td(?<attrs>[^>]*?)\s+title=""(?<title>[^""]*)""(?<rest>[^>]*?)>(?<inner>.*?)</td>";
                htmlContent = Regex.Replace(htmlContent, pattern, match =>
                {
                    string innerText = match.Groups["inner"].Value;
                    // If the cell text is shorter than the threshold, remove the title attribute
                    if (innerText.Length < shortTextThreshold)
                    {
                        // Rebuild the <td> tag without the title attribute
                        string attrs = match.Groups["attrs"].Value;
                        string rest = match.Groups["rest"].Value;
                        return $"<td{attrs}{rest}>{innerText}</td>";
                    }
                    // Keep the original <td> element unchanged
                    return match.Value;
                }, RegexOptions.Singleline | RegexOptions.IgnoreCase);

                File.WriteAllText(outputPath, htmlContent);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
