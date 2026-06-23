using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class ExportNamedRangesToHtml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "NamedRanges.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve all named ranges (and tables) defined in the workbook
            Aspose.Cells.Range[] namedRanges = workbook.Worksheets.GetNamedRangesAndTables();

            // Build a single HTML document
            StringBuilder htmlBuilder = new StringBuilder();
            htmlBuilder.AppendLine("<html><head><title>Named Ranges Documentation</title></head><body>");
            htmlBuilder.AppendLine("<h1>Named Ranges</h1>");

            // Convert each range to HTML and append to the document
            foreach (Aspose.Cells.Range range in namedRanges)
            {
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    ExportNamedRangeAnchors = false // keep output clean
                };

                byte[] htmlBytes = range.ToHtml(options);
                string rangeHtml = Encoding.UTF8.GetString(htmlBytes);

                htmlBuilder.AppendLine($"<h2>{range.Name}</h2>");
                htmlBuilder.AppendLine($"<p>Address: {range.Address}</p>");
                htmlBuilder.AppendLine(rangeHtml);
            }

            htmlBuilder.AppendLine("</body></html>");

            // Save the combined HTML documentation
            File.WriteAllText(outputPath, htmlBuilder.ToString(), Encoding.UTF8);
            Console.WriteLine($"Named ranges exported successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}