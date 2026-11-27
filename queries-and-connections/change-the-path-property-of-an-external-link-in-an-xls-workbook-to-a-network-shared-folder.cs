// Title: Update external link DataSource to a network UNC path in an XLS workbook using Aspose.Cells for .NET
// Common Searches: how to change the DataSource of external links in an XLS file with Aspose.Cells C# | set network UNC path for Excel external links in a .xls workbook using Aspose.Cells | adjust workbook AbsolutePath for relative external references in C# Aspose.Cells | save modified XLS workbook after updating external link paths with Aspose.Cells | C# code to iterate ExternalLinkCollection and modify link paths in an Excel 97-2003 file
// Tags: external link DataSource update Aspose.Cells | network UNC path for Excel external links C# | workbook AbsolutePath configuration Aspose.Cells | save XLS workbook after external link modification | iterate ExternalLinkCollection to change paths

using System;
using System.IO;
using Aspose.Cells;

namespace ChangeExternalLinkPath
{
    // The example loads an XLS workbook, enumerates its ExternalLinkCollection, replaces each link's DataSource with a specified network UNC path, sets the workbook's AbsolutePath to the UNC folder for relative reference resolution, ensures the output directory exists, and saves the updated workbook in Excel97To2003 format using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Validate arguments
                if (args.Length != 3)
                {
                    Console.WriteLine("Usage: ChangeExternalLinkPath <sourceWorkbookPath> <newNetworkPath> <outputWorkbookPath>");
                    return;
                }

                string sourceWorkbookPath = args[0];
                string newNetworkPath = args[1];
                string outputWorkbookPath = args[2];

                // Verify source workbook exists
                if (!File.Exists(sourceWorkbookPath))
                {
                    Console.WriteLine($"Error: Source workbook not found at \"{sourceWorkbookPath}\"");
                    return;
                }

                // Load the workbook (XLS format is automatically detected)
                Workbook workbook = new Workbook(sourceWorkbookPath);

                // Access the collection of external links
                ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

                if (externalLinks == null || externalLinks.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any external links.");
                }
                else
                {
                    Console.WriteLine($"Found {externalLinks.Count} external link(s). Updating their paths...");

                    for (int i = 0; i < externalLinks.Count; i++)
                    {
                        // Update the external link path to the provided network UNC path.
                        externalLinks[i].DataSource = newNetworkPath;
                        Console.WriteLine($"External link #{i + 1} path set to \"{newNetworkPath}\"");
                    }
                }

                // Set the workbook's AbsolutePath to help resolve relative references, if any.
                string networkFolder = Path.GetDirectoryName(newNetworkPath);
                if (!string.IsNullOrEmpty(networkFolder))
                {
                    workbook.AbsolutePath = networkFolder;
                    Console.WriteLine($"Workbook.AbsolutePath set to \"{networkFolder}\"");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputWorkbookPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook. Use Excel97To2003 for XLS format.
                workbook.Save(outputWorkbookPath, SaveFormat.Excel97To2003);
                Console.WriteLine($"Workbook saved successfully to \"{outputWorkbookPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
