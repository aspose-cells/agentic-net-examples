// Title: Set a relative Path for external links in an Excel workbook and verify portability using Aspose.Cells for .NET
// AI Prompts: Iterate over the workbook’s external links, assign each Link.Path to "Data\LinkedWorkbook.xlsx", and save the file as a portable version. | Load an existing workbook, use reflection to locate the ExternalLinks collection, modify the Path property to a relative location, then reload the saved workbook to print the updated paths. | Create a placeholder workbook if it does not exist, adjust any external link paths to a relative folder, save the result, and confirm the changes by reading back the Path values.
// Common Searches: Aspose.Cells how to set external link path to a relative folder in C# | verify that external links remain functional after saving workbook with Aspose.Cells | make Excel workbook portable by changing linked file paths using Aspose.Cells .NET | using reflection to modify external link Path property in Aspose.Cells workbook
// Tags: Aspose.Cells external link path adjustment | C# relative location for linked Excel workbooks | reflection-based ExternalLinks modification | portable Excel workbook with linked files | validate linked file references after saving workbook

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace ExternalLinkPathExampleApp
{
    // The example loads InputWorkbook.xlsx (creating a simple placeholder if missing), uses reflection to access any external links on the first worksheet, sets each link's Path to the relative location "Data\LinkedWorkbook.xlsx", saves the modified workbook as PortableWorkbook.xlsx, then reloads the saved file and prints the updated link paths or reports if the external‑link API is unavailable.
    class ExternalLinkPathExample
    {
        static void Main()
        {
            try
            {
                const string inputFile = "InputWorkbook.xlsx";
                const string portableFile = "PortableWorkbook.xlsx";

                // Ensure the input workbook exists; create a simple placeholder if missing
                if (!File.Exists(inputFile))
                {
                    var placeholder = new Workbook();
                    placeholder.Worksheets[0].Name = "Sheet1";
                    placeholder.Save(inputFile);
                }

                // Load the workbook (may contain external links)
                var workbook = new Workbook(inputFile);

                // Attempt to process external links via reflection (API may be unavailable)
                try
                {
                    var ws = workbook.Worksheets[0];
                    var externalLinksProp = ws.GetType().GetProperty("ExternalLinks");
                    if (externalLinksProp != null)
                    {
                        var externalLinks = externalLinksProp.GetValue(ws) as IEnumerable;
                        if (externalLinks != null)
                        {
                            foreach (var link in externalLinks)
                            {
                                var pathProp = link.GetType().GetProperty("Path");
                                if (pathProp != null && pathProp.CanWrite)
                                {
                                    pathProp.SetValue(link, Path.Combine("Data", "LinkedWorkbook.xlsx"));
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning while processing external links: {ex.Message}");
                }

                // Save the modified workbook
                workbook.Save(portableFile);

                // Reload the saved workbook to verify external link handling (if supported)
                var testWorkbook = new Workbook(portableFile);

                Console.WriteLine("External link paths after saving:");
                try
                {
                    var ws = testWorkbook.Worksheets[0];
                    var externalLinksProp = ws.GetType().GetProperty("ExternalLinks");
                    if (externalLinksProp != null)
                    {
                        var externalLinks = externalLinksProp.GetValue(ws) as IEnumerable;
                        if (externalLinks != null)
                        {
                            foreach (var link in externalLinks)
                            {
                                var pathProp = link.GetType().GetProperty("Path");
                                if (pathProp != null && pathProp.CanRead)
                                {
                                    var pathValue = pathProp.GetValue(link) as string;
                                    Console.WriteLine(pathValue);
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No external links found.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("External link API not available in this Aspose.Cells version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error while reading external links: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
