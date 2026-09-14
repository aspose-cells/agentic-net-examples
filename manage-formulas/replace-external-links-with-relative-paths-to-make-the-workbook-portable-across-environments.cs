// Title: Convert absolute external link paths to relative paths in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel file using Aspose.Cells, iterates through Worksheets.ExternalLinks, and rewrites any absolute LinkPath values to paths relative to the workbook's folder. | Create a C# helper method that receives a base directory and an absolute file path, returns the relative path, and apply this method to each external link before saving the workbook.
// Common Searches: how to replace absolute external link paths with relative ones in Aspose.Cells C# | make Excel workbook portable by converting external link paths using Aspose.Cells | Aspose.Cells C# get relative path for external links in a workbook | convert external link file paths to relative in .NET Excel processing | using reflection to modify external link paths in Aspose.Cells workbook
// Tags: Aspose.Cells external link relative path conversion | C# compute relative file path for Excel workbook | update external link paths via reflection Aspose.Cells | Excel workbook portability Aspose.Cells .NET | convert absolute to relative paths in Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, determines its directory, iterates over all external links via reflection to access the LinkPath property, converts any rooted (absolute) paths to relative paths based on the workbook location using a custom GetRelativePath method, writes the new relative paths back to the links, and saves the modified workbook to a new file, enabling portable workbooks across environments.
class WorkbookLinkRelativizer
{
    static void Main()
    {
        try
        {
            // Path to the original workbook (absolute path)
            string workbookPath = @"C:\Data\MyWorkbook.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Source workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(workbookPath);

            // Directory of the workbook – used as the base for relative paths
            string workbookDirectory = Path.GetDirectoryName(workbookPath);

            // Guard against null directory (should not happen for a valid path)
            if (string.IsNullOrEmpty(workbookDirectory))
            {
                Console.WriteLine("Unable to determine workbook directory.");
                return;
            }

            // Iterate through all external links in the workbook
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                // Use reflection to obtain the current link path (property name may vary between versions)
                var linkPathProp = link.GetType().GetProperty("LinkPath");
                if (linkPathProp == null)
                {
                    Console.WriteLine("ExternalLink does not expose a 'LinkPath' property in this Aspose.Cells version.");
                    continue;
                }

                string currentPath = linkPathProp.GetValue(link) as string;
                if (string.IsNullOrEmpty(currentPath))
                    continue;

                // Process only if the path is rooted (i.e., absolute)
                if (Path.IsPathRooted(currentPath))
                {
                    // Convert the absolute path to a relative one based on the workbook location
                    string relativePath = GetRelativePath(workbookDirectory, currentPath);

                    // Assign the relative path back to the external link via reflection
                    linkPathProp.SetValue(link, relativePath);
                }
            }

            // Prepare output path
            string outputPath = @"C:\Data\MyWorkbook_Portable.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Converts an absolute path to a relative path based on a base folder
    private static string GetRelativePath(string baseFolder, string targetPath)
    {
        // Ensure the base folder ends with a directory separator
        Uri baseUri = new Uri(AppendDirectorySeparatorChar(baseFolder));
        Uri targetUri = new Uri(targetPath);

        // Compute the relative URI and convert it back to a file system path
        Uri relativeUri = baseUri.MakeRelativeUri(targetUri);
        string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

        // Replace URI separators with the OS-specific directory separator
        return relativePath.Replace('/', Path.DirectorySeparatorChar);
    }

    // Helper to guarantee a trailing directory separator on a folder path
    private static string AppendDirectorySeparatorChar(string path)
    {
        if (!path.EndsWith(Path.DirectorySeparatorChar.ToString()))
            return path + Path.DirectorySeparatorChar;
        return path;
    }
}
