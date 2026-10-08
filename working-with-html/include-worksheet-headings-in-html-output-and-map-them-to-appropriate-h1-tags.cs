// Title: Convert the first row of an exported Excel worksheet to <h1> headings in a single HTML file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, saves the active worksheet as a single UTF‑8 HTML file with embedded Base64 images, and then replaces the cells of the first table row with <h1> tags. | Show how to configure Aspose.Cells HtmlSaveOptions to export only the active sheet, include column and row headings, and embed images as Base64 for easy post‑processing of the HTML output. | Write a C# regex routine that transforms the first <tr> element in the HTML produced by Aspose.Cells, converting each <td> or <th> element into an <h1> element.
// Common Searches: how to map Excel header row to h1 tags in HTML using Aspose.Cells C# | Aspose.Cells export active worksheet to single HTML file with embedded images | replace first table row cells with h1 after saving workbook as HTML in .NET | C# regex to change td elements to h1 in Aspose.Cells generated HTML
// Tags: Aspose.Cells HTML export custom heading mapping | C# regex post‑process Aspose.Cells HTML | HtmlSaveOptions export active worksheet as single file | embed images as base64 in Aspose.Cells HTML output | first row cell conversion to h1 in Aspose.Cells HTML

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlHeadingExample
{
    // The example loads an Excel workbook, uses Aspose.Cells HtmlSaveOptions to export the first worksheet as a single UTF‑8 HTML file with column/row headings and Base64‑encoded images, then applies a regex to the generated HTML to replace the cells of the first <tr> with <h1> tags before writing the final file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            string workbookPath = @"C:\Input\Sample.xlsx";
            Workbook workbook = new Workbook(workbookPath);

            // Get the first worksheet (you can adjust the index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export column and row headings (A, B, 1, 2, ...) if needed
                ExportHeadings = true,
                // Export only the active sheet to keep the output simple
                ExportActiveWorksheetOnly = true,
                // Use a single HTML file (no separate CSS or images)
                ExportImagesAsBase64 = true,
                // Set the encoding to UTF-8
                Encoding = Encoding.UTF8
            };

            // Save the workbook to a memory stream first
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, saveOptions);
                htmlStream.Position = 0;

                // Read the generated HTML into a string
                string htmlContent = new StreamReader(htmlStream, Encoding.UTF8).ReadToEnd();

                // Identify the first row of the worksheet in the HTML table.
                // Aspose.Cells renders each row as <tr>...</tr>. We'll replace the first <tr> cells with <h1>.
                // This simple regex assumes the first <tr> corresponds to the heading row.
                string pattern = @"<tr>(.*?)</tr>";
                MatchCollection rows = Regex.Matches(htmlContent, pattern, RegexOptions.Singleline);

                if (rows.Count > 0)
                {
                    // Process the first row (index 0)
                    string firstRow = rows[0].Value;

                    // Replace each <td>...</td> (or <th>...</th>) with <h1>...</h1>
                    string transformedRow = Regex.Replace(firstRow,
                        @"<(td|th)[^>]*>(.*?)</\1>",
                        m => $"<h1>{m.Groups[2].Value}</h1>",
                        RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    // Rebuild the HTML with the transformed first row
                    htmlContent = htmlContent.Replace(firstRow, transformedRow);
                }

                // Write the final HTML to a file (replace with your desired output path)
                string outputPath = @"C:\Output\SampleWithHeadings.html";
                File.WriteAllText(outputPath, htmlContent, Encoding.UTF8);
            }

            Console.WriteLine("HTML export completed with headings mapped to <h1> tags.");
        }
    }
}
