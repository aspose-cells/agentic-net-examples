// Title: Generate a plain‑text report of hidden external link paths in an Excel workbook using Aspose.Cells and C#
// AI Prompts: Load an .xlsx file with Aspose.Cells, iterate each worksheet's ExternalLinks via reflection, filter for links where IsHidden is true, and collect their ExternalLinkPath values. | Create a timestamped header and write the collected hidden link paths to a .txt file, handling missing input files and write‑error exceptions. | Output a console message showing the full path of the generated report and log any sheet‑level processing errors.
// Common Searches: how to list hidden external links in an Excel file using Aspose.Cells C# | C# generate text file report of hidden external link paths from workbook | Aspose.Cells reflection external links IsHidden property example | detect hidden external links in .xlsx with .NET | export hidden external link paths to txt using Aspose.Cells
// Tags: Aspose.Cells hidden external link extraction | C# enumerate worksheet external links via reflection | generate plain text report from Excel workbook | detect hidden external links in .xlsx file | write hidden link paths to text file C#

using System;
using System.IO;
using System.Text;
using System.Collections;
using Aspose.Cells;

namespace HiddenExternalLinksReport
{
    // The program loads an Excel workbook with Aspose.Cells, uses reflection to walk through each worksheet's ExternalLinks collection, selects links marked as hidden (IsHidden = true), extracts their ExternalLinkPath, and writes a timestamped plain‑text report listing each hidden link or indicating none were found, with robust error handling for missing files and write failures.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be analyzed
            string excelPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(excelPath)}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook using Aspose.Cells
                workbook = new Workbook(excelPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // StringBuilder to accumulate the report content
            StringBuilder reportBuilder = new StringBuilder();

            // Header for the report
            reportBuilder.AppendLine("Hidden External Link Paths Report");
            reportBuilder.AppendLine($"Generated on: {DateTime.Now}");
            reportBuilder.AppendLine(new string('=', 40));
            reportBuilder.AppendLine();

            bool hiddenLinkFound = false;

            // Iterate through all worksheets and their external links using reflection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    var externalLinksProp = sheet.GetType().GetProperty("ExternalLinks");
                    if (externalLinksProp == null) continue;

                    var externalLinks = externalLinksProp.GetValue(sheet) as IEnumerable;
                    if (externalLinks == null) continue;

                    foreach (var linkObj in externalLinks)
                    {
                        if (linkObj == null) continue;

                        var isHiddenProp = linkObj.GetType().GetProperty("IsHidden");
                        bool isHidden = isHiddenProp != null && (bool)isHiddenProp.GetValue(linkObj);

                        if (isHidden)
                        {
                            hiddenLinkFound = true;
                            var pathProp = linkObj.GetType().GetProperty("ExternalLinkPath");
                            string linkPath = pathProp != null ? pathProp.GetValue(linkObj)?.ToString() : "N/A";

                            reportBuilder.AppendLine($"Hidden Link (Sheet: {sheet.Name}): {linkPath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // If no hidden links were found, note it in the report
            if (!hiddenLinkFound)
            {
                reportBuilder.AppendLine("No hidden external links were found.");
            }

            // Define the output text file path
            string reportPath = "HiddenExternalLinksReport.txt";

            try
            {
                // Write the report to a plain‑text file
                File.WriteAllText(reportPath, reportBuilder.ToString());
                Console.WriteLine($"Report generated successfully at: {Path.GetFullPath(reportPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report: {ex.Message}");
            }
        }
    }
}
