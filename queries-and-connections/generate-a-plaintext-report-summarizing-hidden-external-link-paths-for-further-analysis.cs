// Title: Create a plain‑text report of hidden external links in an Excel workbook with Aspose.Cells for .NET
// Description: Loads an Excel file using Aspose.Cells, reads the Worksheets.ExternalLinks collection, filters links where IsVisible is false, and generates a formatted .txt report that includes the workbook name, generation time, DataSource, OriginalDataSource, PathType, and IsReferred values.
// Keywords: Aspose.Cells hidden external links | C# list invisible external links | Excel external link report .NET | ExternalLink IsVisible property | generate text report Aspose.Cells | audit hidden data connections Excel
// Common Searches: list hidden external links Aspose.Cells C# | export invisible Excel links to text file | how to detect hidden external connections in .xlsx | Aspose.Cells ExternalLink IsVisible example | generate external link audit report .NET
// Developer Intent: Produce a plain‑text audit file that enumerates all non‑visible external links in a specified workbook.
// Use Cases: Validate workbooks before distribution to ensure no undisclosed external data sources. | Create compliance documentation that captures all hidden connections for regulatory review. | Automate batch scanning of multiple Excel files to log hidden external links for quality control.
// AI Prompts: Write C# code with Aspose.Cells that extracts hidden external links and saves them to a CSV file. | Explain the role of the IsVisible property on ExternalLink objects and how to filter for hidden links. | Refactor the sample to log hidden link details using a logging framework instead of writing a text file.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExternalLinkReport
{
    // Loads an Excel file using Aspose.Cells, reads the Worksheets.ExternalLinks collection, filters links where IsVisible is false, and generates a formatted .txt report that includes the workbook name, generation time, DataSource, OriginalDataSource, PathType, and IsReferred values.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that may contain external links
            string workbookPath = "input.xlsx";

            // Load the workbook (create/load lifecycle)
            Workbook workbook = new Workbook(workbookPath);

            // Get the collection of external links from the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

            // Prepare a StringBuilder to accumulate the plain‑text report
            StringBuilder reportBuilder = new StringBuilder();

            reportBuilder.AppendLine("Hidden External Link Paths Report");
            reportBuilder.AppendLine($"Workbook: {Path.GetFileName(workbookPath)}");
            reportBuilder.AppendLine($"Generated on: {DateTime.Now}");
            reportBuilder.AppendLine(new string('=', 40));

            // Iterate through all external links and list those that are not visible (hidden)
            bool anyHidden = false;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // IsVisible indicates whether the link is shown in Excel; false means hidden
                if (!link.IsVisible)
                {
                    anyHidden = true;
                    reportBuilder.AppendLine($"Link #{i + 1}");
                    reportBuilder.AppendLine($"  Data Source : {link.DataSource}");
                    reportBuilder.AppendLine($"  Original Data Source : {link.OriginalDataSource}");
                    reportBuilder.AppendLine($"  Path Type   : {link.PathType}");
                    reportBuilder.AppendLine($"  Is Referred : {link.IsReferred}");
                    reportBuilder.AppendLine();
                }
            }

            if (!anyHidden)
            {
                reportBuilder.AppendLine("No hidden external links were found in the workbook.");
            }

            // Define the output report file path
            string reportPath = "HiddenExternalLinksReport.txt";

            // Write the report to a plain‑text file (save lifecycle)
            File.WriteAllText(reportPath, reportBuilder.ToString());

            // Optionally, display a confirmation message
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
