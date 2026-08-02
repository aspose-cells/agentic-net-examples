using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExternalLinkReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook path (first argument) or default
            string workbookPath = args.Length > 0 ? args[0] : "input.xlsx";

            // Output report path (second argument) or default
            string reportPath = args.Length > 1 ? args[1] : "HiddenExternalLinksReport.txt";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Prepare a StringBuilder for the plain‑text report
            StringBuilder reportBuilder = new StringBuilder();

            reportBuilder.AppendLine("Hidden External Links Report");
            reportBuilder.AppendLine($"Workbook: {workbookPath}");
            reportBuilder.AppendLine($"Generated: {DateTime.Now}");
            reportBuilder.AppendLine();

            bool foundHidden = false;

            // Iterate through all external links and collect those that are not visible
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                // IsVisible == false indicates a hidden external link
                if (!link.IsVisible)
                {
                    foundHidden = true;
                    reportBuilder.AppendLine($"- DataSource          : {link.DataSource}");
                    reportBuilder.AppendLine($"  OriginalDataSource  : {link.OriginalDataSource}");
                    reportBuilder.AppendLine($"  PathType            : {link.PathType}");
                    reportBuilder.AppendLine();
                }
            }

            if (!foundHidden)
            {
                reportBuilder.AppendLine("No hidden external links were found in the workbook.");
            }

            // Write the report to a plain‑text file (save rule)
            File.WriteAllText(reportPath, reportBuilder.ToString());

            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}