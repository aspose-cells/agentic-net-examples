// Title: Generate a plain‑text report of hidden external links in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to scan a workbook’s ExternalLinkCollection, filter links where IsVisible is false, and output each link’s DataSource, OriginalDataSource, and PathType to a .txt file. | Create a reusable C# method that returns a List of objects representing hidden external links (including index, DataSource, OriginalDataSource, and PathType) from an Excel file with Aspose.Cells. | Enhance the hidden external links report to prepend the workbook name and generation timestamp before listing the link details.
// Common Searches: aspocells c# how to list hidden external links in an xlsx file | export invisible external link details from Excel to txt using Aspose.Cells | C# code to detect external links with IsVisible false in a workbook | generate plain text report of external link metadata with Aspose.Cells .NET
// Tags: Aspose.Cells extract hidden external links | C# write external link metadata to txt | ExternalLinkCollection filter IsVisible false | Excel workbook hidden link report generation | Aspose.Cells external link path type extraction

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates its ExternalLinkCollection, identifies links where IsVisible is false, and writes the link index, DataSource, OriginalDataSource, and PathType to a plain‑text report.
class HiddenExternalLinksReport
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Path for the plain‑text report
        string reportPath = "HiddenExternalLinksReport.txt";

        // Create the report file
        using (StreamWriter writer = new StreamWriter(reportPath))
        {
            writer.WriteLine("Hidden External Links Report");
            writer.WriteLine($"Workbook: {inputPath}");
            writer.WriteLine($"Generated on: {DateTime.Now}");
            writer.WriteLine();

            bool foundHidden = false;

            // Iterate through all external links in the workbook
            ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink link = externalLinks[i];

                // External links with IsVisible == false are considered hidden
                if (!link.IsVisible)
                {
                    foundHidden = true;
                    writer.WriteLine($"Link #{i}");
                    writer.WriteLine($"DataSource: {link.DataSource}");
                    writer.WriteLine($"OriginalDataSource: {link.OriginalDataSource}");
                    writer.WriteLine($"PathType: {link.PathType}");
                    writer.WriteLine();
                }
            }

            if (!foundHidden)
            {
                writer.WriteLine("No hidden external links were found in this workbook.");
            }
        }

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
