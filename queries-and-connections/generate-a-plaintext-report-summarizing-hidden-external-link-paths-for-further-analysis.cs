// Title: Generate a plain‑text report of hidden external link paths from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells retrieve hidden external link paths from an .xlsx file | how to list external links in Excel using Aspose.Cells reflection | generate a text report of worksheet external links with Aspose.Cells for .NET | extract SourceFullName of hidden external links from Excel workbook in C# | use reflection to access Worksheet ExternalLinks collection in Aspose.Cells
// Tags: Aspose.Cells external links extraction | C# reflection worksheet external links | plain text Excel link report generation | hidden external link path retrieval | write external links report to file C#

using System;
using System.Collections;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExternalLinkReport
{
    // The sample loads an Excel workbook with Aspose.Cells, checks file existence, and uses reflection to obtain each worksheet's ExternalLinks collection. It iterates through the links, safely reads the SourceFullName and optional Name, appends the information to a StringBuilder, and writes a plain‑text report (ExternalLinksReport.txt). The code handles missing files, unavailable properties, and write errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Prepare a StringBuilder for the report
            StringBuilder reportBuilder = new StringBuilder();

            // Report header
            reportBuilder.AppendLine("Hidden External Link Paths Report");
            reportBuilder.AppendLine($"Generated on: {DateTime.Now}");
            reportBuilder.AppendLine(new string('=', 40));
            reportBuilder.AppendLine();

            int linkIndex = 0;

            // Iterate through all worksheets and attempt to retrieve external links via reflection/dynamic
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                IEnumerable externalLinks = GetExternalLinksCollection(sheet);
                if (externalLinks == null) continue;

                foreach (object linkObj in externalLinks)
                {
                    dynamic link = linkObj;
                    linkIndex++;
                    reportBuilder.AppendLine($"Link #{linkIndex}");

                    // Attempt to read the path (SourceFullName) safely
                    try
                    {
                        string path = link.SourceFullName as string;
                        reportBuilder.AppendLine($"Path: {path}");
                    }
                    catch
                    {
                        reportBuilder.AppendLine("Path: N/A");
                    }

                    // Attempt to read the link name safely
                    try
                    {
                        string name = link.Name as string;
                        if (!string.IsNullOrEmpty(name))
                        {
                            reportBuilder.AppendLine($"Name: {name}");
                        }
                    }
                    catch
                    {
                        // Ignore if Name property is unavailable
                    }

                    reportBuilder.AppendLine(new string('-', 30));
                }
            }

            // If no external links were found, note it in the report
            if (linkIndex == 0)
            {
                reportBuilder.AppendLine("No external links were found in the workbook.");
            }

            // Output report file path
            string outputPath = "ExternalLinksReport.txt";

            try
            {
                // Write the report to a text file
                File.WriteAllText(outputPath, reportBuilder.ToString());
                Console.WriteLine($"Report generated successfully at: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write report: {ex.Message}");
            }
        }

        // Retrieves the ExternalLinks collection from a Worksheet using reflection.
        // Returns null if the property does not exist or cannot be accessed.
        private static IEnumerable GetExternalLinksCollection(Worksheet sheet)
        {
            try
            {
                var propInfo = sheet.GetType().GetProperty("ExternalLinks");
                if (propInfo != null)
                {
                    var value = propInfo.GetValue(sheet) as IEnumerable;
                    return value;
                }
            }
            catch
            {
                // Ignored – property not available in this version.
            }
            return null;
        }
    }
}
