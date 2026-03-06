using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be examined
        string inputPath = "input.xlsx";

        // Path where the text report will be saved
        string outputPath = "ExternalLinksReport.txt";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Retrieve the collection of external links in the workbook
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Build a textual report
        StringBuilder reportBuilder = new StringBuilder();

        if (externalLinks.Count == 0)
        {
            Console.WriteLine("Workbook contains no external links.");
            reportBuilder.AppendLine("Workbook contains no external links.");
        }
        else
        {
            Console.WriteLine($"Workbook contains {externalLinks.Count} external link(s).");
            reportBuilder.AppendLine($"Workbook contains {externalLinks.Count} external link(s).");

            // Iterate through each external link and report its visibility
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];
                string visibility = link.IsVisible ? "Visible" : "Hidden";
                string line = $"Link {i + 1}: DataSource = {link.DataSource}, Visibility = {visibility}";
                Console.WriteLine(line);
                reportBuilder.AppendLine(line);
            }
        }

        // Write the report to a text file
        File.WriteAllText(outputPath, reportBuilder.ToString());

        Console.WriteLine($"Report written to {outputPath}");
    }
}