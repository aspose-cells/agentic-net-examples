// Title: Check external source file existence before updating Excel ExternalLink.SourceFilePath with Aspose.Cells for .NET
// Common Searches: how to verify external link file exists before modifying it with Aspose.Cells in C# | Aspose.Cells update ExternalLink source path only when the source file is present | C# code to confirm external workbook file before setting SourceFilePath using Aspose.Cells | using reflection to change ExternalLink.SourceFilePath after confirming file existence in Aspose.Cells | prevent errors when updating Excel external links with missing source files in .NET
// Tags: external link file existence validation Aspose.Cells | set ExternalLink SourceFilePath via reflection C# | conditional Excel external link update .NET | Aspose.Cells workbook external link modification | avoid missing source file errors Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example demonstrates how to verify that both the target workbook and the new external source file exist, then load the workbook with Aspose.Cells, iterate through its ExternalLink collection, use reflection to assign a new SourceFilePath when the property is writable, and finally save the workbook with the updated link.
class ExternalLinkValidator
{
    static void Main()
    {
        // Paths
        string workbookPath = @"C:\Docs\MyWorkbook.xlsx";
        string newExternalLinkPath = @"C:\Data\SourceFile.xlsx";

        // Validate that the external link file exists
        if (!File.Exists(newExternalLinkPath))
        {
            Console.WriteLine($"Error: The external file \"{newExternalLinkPath}\" does not exist.");
            return;
        }

        // Validate that the workbook file exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The workbook file \"{workbookPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Iterate through all external links in the workbook
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                try
                {
                    // Use reflection to set the source file path if the property exists
                    PropertyInfo prop = link.GetType().GetProperty("SourceFilePath", BindingFlags.Public | BindingFlags.Instance);
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(link, newExternalLinkPath);
                        Console.WriteLine($"Updated external link to \"{newExternalLinkPath}\".");
                    }
                    else
                    {
                        Console.WriteLine("The ExternalLink object does not expose a writable SourceFilePath property.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to update an external link: {ex.Message}");
                }
            }

            // Save the workbook with the updated external link
            string outputPath = @"C:\Docs\MyWorkbook_Updated.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\" with updated external link.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
