// Title: Convert absolute external reference file names to relative paths in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change external link file name to relative path | how to make an Excel workbook portable by updating external references with Aspose.Cells | C# compute relative path for external reference in an Aspose.Cells workbook | replace absolute external reference paths in .xlsx using Aspose.Cells API | Aspose.Cells .NET update ExternalReference.FileName to a relative path
// Tags: Aspose.Cells external reference relative path conversion | C# compute relative file path for workbook | update ExternalReference.FileName property Aspose.Cells | portable Excel workbook external links | reflection based external links handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections;
using System.IO;

// The example loads an Excel file with Aspose.Cells, iterates each worksheet's external links via reflection, identifies external references, computes a relative path based on the workbook's directory, replaces the absolute FileName with the relative path, and saves the workbook, making the file portable across environments.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Directory of the workbook (used for relative path calculation)
            string workbookDirectory = Path.GetDirectoryName(Path.GetFullPath(inputPath)) ?? string.Empty;

            // Iterate through all worksheets and their external links using reflection (to stay compatible with different Aspose.Cells versions)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var externalLinksProp = sheet.GetType().GetProperty("ExternalLinks");
                if (externalLinksProp == null) continue; // No external links support in this version

                var externalLinksObj = externalLinksProp.GetValue(sheet);
                if (externalLinksObj is not IEnumerable externalLinks) continue;

                foreach (var externalLink in externalLinks)
                {
                    try
                    {
                        var linkType = externalLink.GetType();

                        // Check if the link is an external reference
                        var isExternalProp = linkType.GetProperty("IsExternalReference");
                        bool isExternal = isExternalProp?.GetValue(externalLink) as bool? == true;
                        if (!isExternal) continue;

                        // Get the ExternalReference object
                        var externalRefProp = linkType.GetProperty("ExternalReference");
                        var externalReference = externalRefProp?.GetValue(externalLink);
                        if (externalReference == null) continue;

                        // Get the absolute file name
                        var fileNameProp = externalReference.GetType().GetProperty("FileName");
                        string absolutePath = fileNameProp?.GetValue(externalReference) as string;
                        if (string.IsNullOrEmpty(absolutePath)) continue;

                        // Compute relative path and replace
                        string relativePath = GetRelativePath(absolutePath, workbookDirectory);
                        fileNameProp?.SetValue(externalReference, relativePath);
                    }
                    catch (Exception exLink)
                    {
                        Console.WriteLine($"Warning: Unable to process an external link. {exLink.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to compute a relative path between two absolute paths
    private static string GetRelativePath(string targetPath, string basePath)
    {
        if (string.IsNullOrEmpty(targetPath) || string.IsNullOrEmpty(basePath))
            return targetPath ?? string.Empty;

        // Ensure base path ends with a directory separator
        if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            basePath += Path.DirectorySeparatorChar;

        Uri baseUri = new Uri(basePath);
        Uri targetUri = new Uri(targetPath);

        // Build the relative URI and convert it back to a file system path
        Uri relativeUri = baseUri.MakeRelativeUri(targetUri);
        string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

        // Convert URI separators to the OS-specific directory separator
        return relativePath.Replace('/', Path.DirectorySeparatorChar);
    }
}
