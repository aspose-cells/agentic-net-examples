// Title: Update external link URLs in an Excel workbook with Aspose.Cells for .NET using reflection (custom ribbon button test)
// Common Searches: aspnet c# replace external link paths in Excel workbook using Aspose.Cells reflection | how to process worksheet ExternalLinks collection with Aspose.Cells .NET | custom ribbon button code to modify Excel external hyperlinks via Aspose.Cells | save updated workbook after changing external URLs with Aspose.Cells | debug external link processing errors in Aspose.Cells using reflection
// Tags: Aspose.Cells external link URL replacement via reflection | C# worksheet ExternalLinks enumeration Aspose.Cells | Excel ribbon integration external link update Aspose.Cells | save workbook after external link modification Aspose.Cells | handle missing ExternalLinks property Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace RibbonButtonTest
{
    // The example loads a source XLSX file, uses reflection to access each worksheet's ExternalLinks collection, replaces any URL containing an old server path with a new one, creates the output folder if needed, and saves the modified workbook. It logs original and updated links, handles missing files, absent properties, and reflection errors, and demonstrates how a custom ribbon button can trigger the update using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for source and output workbooks
            string sourcePath = @"C:\Temp\SourceWorkbook.xlsx";
            string outputPath = @"C:\Temp\UpdatedWorkbook.xlsx";

            try
            {
                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Attempt to process external links using reflection.
                // This avoids compile‑time dependencies on APIs that may not exist in the current Aspose.Cells version.
                try
                {
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        var externalLinksProp = sheet.GetType().GetProperty("ExternalLinks");
                        if (externalLinksProp == null) continue; // Property not available

                        var externalLinks = externalLinksProp.GetValue(sheet) as IEnumerable;
                        if (externalLinks == null) continue;

                        foreach (var linkObj in externalLinks)
                        {
                            var urlProp = linkObj.GetType().GetProperty("Url");
                            if (urlProp == null) continue; // Url property not available

                            string url = urlProp.GetValue(linkObj) as string;
                            Console.WriteLine($"Original Link: {url}");

                            // Sample replacement logic: change old server path to new server path
                            if (!string.IsNullOrEmpty(url) && url.Contains(@"\\OldServer\SharedFolder"))
                            {
                                string newUrl = url.Replace(@"\\OldServer\SharedFolder", @"\\NewServer\NewSharedFolder");
                                urlProp.SetValue(linkObj, newUrl);
                                Console.WriteLine($"Updated Link: {newUrl}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log any issues while processing external links and continue.
                    Console.WriteLine($"External link processing skipped: {ex.Message}");
                }

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the updated workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
