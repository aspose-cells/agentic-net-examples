using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsHtmlSpaceCollapse
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert text that contains a newline followed by multiple spaces
            // Example: "Line1\n   Line2"
            sheet.Cells["A1"].PutValue("Line1\n   Line2");

            // Configure HTML save options (default options are sufficient)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Define the output HTML file path
            string htmlPath = Path.Combine(Environment.CurrentDirectory, "output.html");

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Collapse multiple spaces that appear immediately after a newline character
            // Pattern explanation:
            //   (?<=\n)   - positive lookbehind to ensure we are after a newline
            //   {2,}     - match two or more space characters
            // Replace with a single space
            string collapsedContent = Regex.Replace(htmlContent, @"(?<=\n) {2,}", " ");

            // Overwrite the HTML file with the cleaned content
            File.WriteAllText(htmlPath, collapsedContent);

            Console.WriteLine($"HTML file saved and spaces collapsed: {htmlPath}");
        }
    }
}