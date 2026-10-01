// Title: Update Excel workbook external links to relative paths using Aspose.Cells for .NET (C#)
// AI Prompts: Load the workbook, iterate workbook.Worksheets.ExternalLinks, and replace each SourceFullName that is an absolute path with a path relative to the workbook folder. | Create Uri objects for the workbook directory and the target file, compute the relative URI, convert it to a file system string, and assign it back to the external link. | Save the workbook after all external links have been rewritten to preserve the new relative references.
// Common Searches: how to make external links in an Excel file portable with Aspose.Cells C# | convert absolute external link paths to relative in .NET workbook | Aspose.Cells update external connection source path to relative | C# change Excel external link SourceFullName to relative path | relative path for external links in Aspose.Cells workbook
// Tags: external link sourcefullName conversion Aspose.Cells | absolute external link to relative Excel .NET | Aspose.Cells external connections path adjustment | portable workbook external references C# | relative URI calculation for Excel links

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, walks through its external links, replaces any absolute SourceFullName values with paths computed relative to the workbook's directory, and saves the file, making the workbook portable across environments.
class UpdateExternalLinks
{
    static void Main()
    {
        // Path to the workbook that contains external links
        string workbookPath = @"C:\Data\MyWorkbook.xlsx";

        // Verify that the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook not found at '{workbookPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Directory of the workbook – used as the base for relative paths
            string workbookDirectory = Path.GetDirectoryName(workbookPath) ?? string.Empty;
            // Ensure the directory ends with a separator for correct Uri handling
            Uri workbookDirUri = new Uri(workbookDirectory + Path.DirectorySeparatorChar);

            // Iterate through all external links in the workbook
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                try
                {
                    // Use dynamic to access SourceFullName (property name may vary across versions)
                    dynamic dLink = link;
                    string originalPath = dLink.SourceFullName as string;

                    // Process only if the path is absolute (rooted)
                    if (!string.IsNullOrEmpty(originalPath) && Path.IsPathRooted(originalPath))
                    {
                        // Create a Uri for the target file
                        Uri targetUri = new Uri(originalPath);

                        // Compute the relative Uri from the workbook directory to the target
                        Uri relativeUri = workbookDirUri.MakeRelativeUri(targetUri);

                        // Convert the relative Uri to a file system path
                        string relativePath = Uri.UnescapeDataString(relativeUri.ToString())
                                                .Replace('/', Path.DirectorySeparatorChar);

                        // Update the link to use the relative path
                        dLink.SourceFullName = relativePath;
                    }
                }
                catch (Exception exLink)
                {
                    Console.WriteLine($"Failed to process a link: {exLink.Message}");
                }
            }

            // Save the workbook
            workbook.Save(workbookPath);
            Console.WriteLine("External links updated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
