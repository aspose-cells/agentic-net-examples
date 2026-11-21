// Title: Detect hidden external links in an Excel workbook and list their source file paths using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# find hidden external links in an .xlsx workbook | list source file paths of hidden external connections with Aspose.Cells .NET | retrieve hidden external link information from worksheets using Aspose.Cells API | C# code to enumerate hidden external links in Excel using Aspose.Cells | detect hidden external links in Excel file via Aspose.Cells reflection
// Tags: hidden external link detection Aspose.Cells | external link source path extraction .NET | worksheet ExternalLinks collection usage C# | reflection access to ExternalLinks Aspose.Cells | enumerate hidden connections in Excel workbook

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

// The sample checks that the input.xlsx file exists, loads it with Aspose.Cells, iterates each worksheet, uses reflection to obtain the ExternalLinks collection, filters links where IsHidden is true, and prints their SourceFilePath, handling missing properties and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use reflection to obtain the ExternalLinks property (may not exist in older versions)
                var externalLinksProp = sheet.GetType().GetProperty("ExternalLinks");
                if (externalLinksProp == null) continue;

                var externalLinksObj = externalLinksProp.GetValue(sheet);
                if (externalLinksObj is IEnumerable externalLinks)
                {
                    foreach (var linkInfo in externalLinks)
                    {
                        // Use reflection to read IsHidden and SourceFilePath properties
                        var isHiddenProp = linkInfo.GetType().GetProperty("IsHidden");
                        var sourcePathProp = linkInfo.GetType().GetProperty("SourceFilePath");

                        if (isHiddenProp != null && sourcePathProp != null)
                        {
                            bool isHidden = (bool)isHiddenProp.GetValue(linkInfo);
                            if (isHidden)
                            {
                                string sourcePath = sourcePathProp.GetValue(linkInfo) as string;
                                Console.WriteLine(sourcePath);
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
