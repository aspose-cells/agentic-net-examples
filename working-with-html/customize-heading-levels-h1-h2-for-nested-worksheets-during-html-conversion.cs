using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlHeadingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Main Sheet";
            sheet1.Cells["A1"].PutValue("Header");
            sheet1.Cells["A2"].PutValue("Data 1");

            Worksheet sheet2 = workbook.Worksheets.Add("Sub Sheet");
            sheet2.Cells["A1"].PutValue("Sub Header");
            sheet2.Cells["A2"].PutValue("Sub Data");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export row/column headings (A, B, 1, 2, etc.)
                ExportRowColumnHeadings = true,

                // Add custom CSS to style the generated <h1> and <h2> tags
                CssStyles = @"
                    h1 {font-size:28px; color:#2E86C1; margin-top:20px;}
                    h2 {font-size:22px; color:#117A65; margin-top:15px;}
                "
            };

            // Save the workbook to a memory stream first
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, htmlOptions);
                ms.Position = 0;

                // Read the generated HTML as text
                string htmlContent = new StreamReader(ms).ReadToEnd();

                // Aspose.Cells uses <h1> for each worksheet name.
                // Replace the second and subsequent <h1> tags with <h2> to create a hierarchy.
                int occurrence = 0;
                htmlContent = System.Text.RegularExpressions.Regex.Replace(
                    htmlContent,
                    @"<h1>(.*?)</h1>",
                    match =>
                    {
                        occurrence++;
                        // First worksheet stays as <h1>, others become <h2>
                        string tag = occurrence == 1 ? "h1" : "h2";
                        return $"<{tag}>{match.Groups[1].Value}</{tag}>";
                    },
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Write the modified HTML to a file
                string outputPath = "NestedWorksheets.html";
                File.WriteAllText(outputPath, htmlContent);
                Console.WriteLine($"HTML file saved with custom heading levels: {outputPath}");
            }
        }
    }
}