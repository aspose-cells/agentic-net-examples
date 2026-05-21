using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExternalLinkReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that needs to be analyzed
            string workbookPath = "InputWorkbook.xlsx";

            // Load the workbook (load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external links from the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Prepare a StringBuilder to compose the plain‑text report
            StringBuilder reportBuilder = new StringBuilder();

            reportBuilder.AppendLine("Hidden External Links Report");
            reportBuilder.AppendLine($"Workbook: {workbookPath}");
            reportBuilder.AppendLine($"Total External Links: {externalLinks.Count}");
            reportBuilder.AppendLine();

            // Iterate through each external link and collect those that are not visible (hidden)
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // IsVisible property indicates whether the link is visible in Excel.
                // Hidden links have IsVisible == false.
                if (!link.IsVisible)
                {
                    // DataSource holds the path of the external link.
                    reportBuilder.AppendLine($"Hidden Link #{i + 1}: {link.DataSource}");
                }
            }

            // If no hidden links were found, note that in the report
            if (reportBuilder.ToString().Contains("Hidden Link #") == false)
            {
                reportBuilder.AppendLine("No hidden external links were found.");
            }

            // Define the output report file path
            string reportPath = "HiddenExternalLinksReport.txt";

            // Write the report to a plain‑text file
            File.WriteAllText(reportPath, reportBuilder.ToString());

            // Optionally, save the workbook (save rule) if any modifications were made.
            // In this scenario we only read data, so saving is not required.
            // workbook.Save("ModifiedWorkbook.xlsx");

            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}