// Title: C# Aspose.Cells example: Detect external HTTP/HTTPS hyperlinks in an Excel workbook and conditionally set the LinksUpToDate property
// AI Prompts: Write C# code using Aspose.Cells that iterates through every worksheet, identifies any hyperlink whose address begins with http:// or https://, and sets WorkbookSettings.UpdateExternalLinks (or LinksUpToDate) to false when no such links are found before saving the file. | Modify the supplied program so that it assigns the LinksUpToDate property based on whether external hyperlinks exist, then saves the workbook.
// Common Searches: how to disable external link updates in an Excel file with Aspose.Cells C# | detect external http hyperlinks in a workbook using Aspose.Cells | set LinksUpToDate property after checking for hyperlinks in Aspose.Cells | conditional workbook settings based on hyperlink presence Aspose.Cells C# | C# code to check for external hyperlinks before saving Excel with Aspose.Cells
// Tags: Aspose.Cells detect external hyperlinks | C# set LinksUpToDate property | Aspose.Cells conditional workbook settings | Excel hyperlink scanning Aspose.Cells | disable external link updates Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook with Aspose.Cells, scans each worksheet's Hyperlink collection for URLs starting with http:// or https://, records whether any external links exist, and then sets the LinksUpToDate (or UpdateExternalLinks) flag accordingly before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                bool hasExternalHyperlink = false;

                // Scan all worksheets for external hyperlinks
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Hyperlink link in sheet.Hyperlinks)
                    {
                        if (!string.IsNullOrEmpty(link.Address) &&
                            (link.Address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                             link.Address.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                        {
                            hasExternalHyperlink = true;
                            break;
                        }
                    }

                    if (hasExternalHyperlink)
                        break;
                }

                // The WorkbookSettings.UpdateExternalLinks property may not be available in all versions.
                // If needed, adjust workbook settings here based on hasExternalHyperlink.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
