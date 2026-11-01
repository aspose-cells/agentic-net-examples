// Title: How to detect and list hidden external links in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, accesses the ExternalLinks collection via reflection, and prints the SourceFullName of each link where IsHidden is true. | Show a robust .NET method that checks for the presence of the ExternalLinks property, safely reads hidden external connections, and returns their source paths as a string list. | Provide a sample program that handles missing files and absent ExternalLinks members while enumerating hidden workbook links with Aspose.Cells.
// Common Searches: aspocells list hidden external links in workbook c# | c# detect hidden external connections in xlsx using Aspose.Cells | how to read ExternalLinks collection with reflection Aspose.Cells older version | enumerate hidden external link source paths .NET Excel file | handle missing ExternalLinks property Aspose.Cells example
// Tags: list hidden external links Aspose.Cells | retrieve ExternalLinks collection via reflection | filter external links by IsHidden property | enumerate workbook external connections .NET | handle absent ExternalLinks property Aspose.Cells

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through each worksheet, uses reflection to obtain the ExternalLinks collection when available, checks the IsHidden flag on each link, and outputs the SourceFullName of hidden links. It includes file‑existence validation and graceful handling of missing properties for compatibility with older Aspose.Cells versions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and attempt to read external links via reflection
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Use reflection to get the ExternalLinks property (may not exist in older versions)
                    var externalLinksProp = sheet.GetType().GetProperty("ExternalLinks");
                    if (externalLinksProp == null)
                    {
                        // Property not available – skip this worksheet
                        continue;
                    }

                    var externalLinks = externalLinksProp.GetValue(sheet) as IEnumerable;
                    if (externalLinks == null) continue;

                    foreach (var linkObj in externalLinks)
                    {
                        // Use reflection to access IsHidden and SourceFullName safely
                        var linkType = linkObj.GetType();

                        var isHiddenProp = linkType.GetProperty("IsHidden");
                        var sourceFullNameProp = linkType.GetProperty("SourceFullName");

                        bool isHidden = isHiddenProp != null && (bool)isHiddenProp.GetValue(linkObj);
                        string sourceFullName = sourceFullNameProp != null ? sourceFullNameProp.GetValue(linkObj)?.ToString() : string.Empty;

                        if (isHidden && !string.IsNullOrEmpty(sourceFullName))
                        {
                            Console.WriteLine(sourceFullName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
