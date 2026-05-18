using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace HiddenExternalLinkReporter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect the first argument to be the path of the workbook to analyze.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the workbook as a command‑line argument.");
                return;
            }

            string workbookPath = args[0];

            // Load the workbook (load rule).
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external links.
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Prepare a StringBuilder to compose the plain‑text report.
            StringBuilder reportBuilder = new StringBuilder();

            reportBuilder.AppendLine("=== Hidden External Links Report ===");
            reportBuilder.AppendLine($"Workbook: {workbookPath}");
            reportBuilder.AppendLine($"Total external links: {externalLinks.Count}");
            reportBuilder.AppendLine();

            // Iterate through each external link and record those that are not visible.
            foreach (ExternalLink link in externalLinks)
            {
                // IsVisible indicates whether the link is shown in Excel.
                // Hidden links have IsVisible == false.
                if (!link.IsVisible)
                {
                    // DataSource holds the path or URL of the external link.
                    reportBuilder.AppendLine($"Hidden Link Path: {link.DataSource}");
                }
            }

            // Define the output report file name.
            string reportPath = "HiddenExternalLinksReport.txt";

            // Write the report to a plain‑text file (standard .NET file I/O).
            File.WriteAllText(reportPath, reportBuilder.ToString());

            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }
    }
}