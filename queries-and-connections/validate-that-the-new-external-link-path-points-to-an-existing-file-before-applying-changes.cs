// Title: Check file existence before updating external link DataSource in an Aspose.Cells workbook (C#)
// AI Prompts: Use System.IO.File.Exists to confirm a new file path is present before assigning it to ExternalLink.DataSource. | Iterate through Workbook.Worksheets.ExternalLinks and set each link's DataSource only when the target file is found. | Save the workbook to a new file after conditionally updating external link paths based on file existence.
// Common Searches: Aspose.Cells C# update external link path only if file exists | validate external link DataSource before saving workbook using Aspose.Cells | C# check file existence before changing external link in Excel with Aspose.Cells | conditional external link update Aspose.Cells workbook example
// Tags: external link datasource validation Aspose.Cells | conditional update of external links C# | file existence check before setting ExternalLink.DataSource | Aspose.Cells workbook external links handling | System.IO.File.Exists with Aspose.Cells external links

using System;
using System.IO;
using Aspose.Cells;

// The code loads an Excel workbook, iterates its external links, verifies that a new external file path exists, updates each link's DataSource to the new path only when the file is found, logs the result, and saves the modified workbook as output.xlsx.
class ValidateExternalLinkPath
{
    static void Main()
    {
        // Path to the workbook that contains external links
        string workbookPath = "input.xlsx";

        // New external file path to set
        string newExternalPath = @"C:\Data\external.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(workbookPath);

        // Get the collection of external links
        ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;

        // Iterate through each external link
        for (int i = 0; i < externalLinks.Count; i++)
        {
            ExternalLink link = externalLinks[i];

            // Validate that the new path points to an existing file
            if (File.Exists(newExternalPath))
            {
                // Apply the new data source
                link.DataSource = newExternalPath;
                Console.WriteLine($"External link at index {i} updated to: {newExternalPath}");
            }
            else
            {
                // Skip updating and inform the user
                Console.WriteLine($"File not found: {newExternalPath}. External link at index {i} not changed.");
            }
        }

        // Save the workbook with updated external links
        workbook.Save("output.xlsx");
    }
}
