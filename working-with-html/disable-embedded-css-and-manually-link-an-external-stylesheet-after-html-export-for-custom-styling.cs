// Title: Export an Excel workbook to HTML with Aspose.Cells in C# while disabling embedded CSS and linking a custom external stylesheet
// Common Searches: Aspose.Cells C# export workbook to HTML without inline CSS | remove embedded <style> tag from Aspose.Cells generated HTML using C# | add external CSS link to HTML saved by Aspose.Cells .NET | disable CSS embedding in Aspose.Cells HtmlSaveOptions C# | customize Aspose.Cells HTML output with a separate stylesheet
// Tags: Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately | C# replace embedded style block with external CSS | link external stylesheet in Aspose.Cells HTML export | custom stylesheet for Aspose.Cells generated HTML | disable embedded CSS in Aspose.Cells HTML output

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlCssDemo
{
    // The example creates a workbook, sets HtmlSaveOptions.ExportWorksheetCSSSeparately to generate a separate CSS file, saves the workbook as HTML, then reads the HTML, removes any <style> elements with a regex, inserts a <link> tag referencing a custom.css file before </head>, writes the modified HTML back, and finally writes the custom CSS file containing table styling for the exported page.
    class Program
    {
        static void Main(string[] args)
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook with some data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Demo";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(2.50);

            // ------------------------------------------------------------
            // 2. Define output folder and HTML file name
            // ------------------------------------------------------------
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "HtmlExport");
            Directory.CreateDirectory(outputFolder);

            string htmlFilePath = Path.Combine(outputFolder, "Workbook.html");
            string externalCssFileName = "custom.css";
            string externalCssPath = Path.Combine(outputFolder, externalCssFileName);

            // ------------------------------------------------------------
            // 3. Configure HtmlSaveOptions to *not* embed CSS
            //    ExportWorksheetCSSSeparately = true creates a separate CSS file,
            //    but we will replace it with our own stylesheet.
            // ------------------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true,   // generate CSS as a separate file
                ExportImagesAsBase64 = false,          // keep images as separate files (optional)
                ExportGridLines = true,                // keep grid lines for visual parity
                // You can set other options here as needed
            };

            // ------------------------------------------------------------
            // 4. Save the workbook as HTML
            // ------------------------------------------------------------
            workbook.Save(htmlFilePath, htmlOptions);
            Console.WriteLine($"Workbook saved as HTML to: {htmlFilePath}");

            // ------------------------------------------------------------
            // 5. Read the generated HTML, remove any embedded <style> blocks,
            //    and insert a <link> tag that points to our external stylesheet.
            // ------------------------------------------------------------
            string htmlContent = File.ReadAllText(htmlFilePath);

            // Remove any <style>...</style> blocks (embedded CSS)
            string stylePattern = @"<style[^>]*?>.*?</style>";
            string cleanedHtml = Regex.Replace(htmlContent, stylePattern,
                                                string.Empty,
                                                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Prepare the <link> tag for the external CSS
            string linkTag = $"<link rel=\"stylesheet\" type=\"text/css\" href=\"{externalCssFileName}\" />{Environment.NewLine}";

            // Insert the link tag just before </head>
            int headCloseIndex = cleanedHtml.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
            if (headCloseIndex >= 0)
            {
                cleanedHtml = cleanedHtml.Insert(headCloseIndex, linkTag);
            }
            else
            {
                // Fallback: prepend to the file if </head> not found
                cleanedHtml = linkTag + cleanedHtml;
            }

            // Write the modified HTML back to the same file
            File.WriteAllText(htmlFilePath, cleanedHtml);
            Console.WriteLine("Embedded CSS removed and external stylesheet linked.");

            // ------------------------------------------------------------
            // 6. Create the external CSS file with custom styling
            // ------------------------------------------------------------
            string customCss = @"
/* Custom stylesheet for the exported HTML */
table {
    border-collapse: collapse;
    width: 100%;
}
th, td {
    border: 1px solid #555;
    padding: 8px;
    text-align: left;
}
th {
    background-color: #4CAF50;
    color: white;
}
tr:nth-child(even) {
    background-color: #f2f2f2;
}
";
            File.WriteAllText(externalCssPath, customCss.Trim());
            Console.WriteLine($"Custom CSS written to: {externalCssPath}");

            // ------------------------------------------------------------
            // 7. Completion message
            // ------------------------------------------------------------
            Console.WriteLine("HTML export with external CSS is complete.");
            Console.WriteLine("Open the HTML file in a browser to see the custom styling applied.");
        }
    }
}
