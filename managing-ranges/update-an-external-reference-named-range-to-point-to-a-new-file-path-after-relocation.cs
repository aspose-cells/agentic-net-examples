using System;
using System.IO;
using Aspose.Cells;

class UpdateExternalReference
{
    static void Main()
    {
        try
        {
            // Input and output workbook paths
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input workbook not found: {inputPath}");

            // Load the workbook that contains the external named range
            Workbook workbook = new Workbook(inputPath);

            // Name of the external named range to be updated
            const string namedRangeName = "MyExternalRange";

            // New folder where the external file has been moved
            const string newFolder = @"D:\NewPath";

            // Retrieve the named range object; ensure it exists
            Name namedRange = workbook.Worksheets.Names[namedRangeName];
            if (namedRange == null)
                throw new InvalidOperationException($"Named range '{namedRangeName}' not found.");

            // Get all areas referred by this name (including external links)
            ReferredArea[] referredAreas = namedRange.GetReferredAreas(false);

            // Iterate through each referred area to locate external links
            foreach (ReferredArea area in referredAreas)
            {
                if (area.IsExternalLink)
                {
                    // The external file name (e.g., "External.xlsx")
                    string externalFileName = area.ExternalFileName;

                    // Build the new full path for the external file
                    string newFullPath = Path.Combine(newFolder, externalFileName);

                    // Update the matching ExternalLink in the workbook's collection
                    ExternalLinkCollection externalLinks = workbook.Worksheets.ExternalLinks;
                    if (externalLinks != null)
                    {
                        for (int i = 0; i < externalLinks.Count; i++)
                        {
                            ExternalLink link = externalLinks[i];
                            if (Path.GetFileName(link.DataSource)
                                    .Equals(externalFileName, StringComparison.OrdinalIgnoreCase))
                            {
                                link.DataSource = newFullPath;
                                link.OriginalDataSource = newFullPath;
                            }
                        }
                    }

                    // Update the named range's RefersTo formula to use the new path
                    string oldRefersTo = namedRange.RefersTo;
                    string updatedRefersTo = oldRefersTo;

                    // Locate the path segment inside single quotes and replace it
                    int firstQuote = oldRefersTo.IndexOf('\'');
                    int exclam = oldRefersTo.IndexOf('!', firstQuote);
                    if (firstQuote >= 0 && exclam > firstQuote)
                    {
                        string oldPathSegment = oldRefersTo.Substring(firstQuote + 1, exclam - firstQuote - 1);
                        updatedRefersTo = oldRefersTo.Replace(oldPathSegment, newFullPath);
                    }
                    else
                    {
                        // Fallback: replace just the file name if the full pattern is not found
                        updatedRefersTo = oldRefersTo.Replace(externalFileName, newFullPath);
                    }

                    namedRange.RefersTo = updatedRefersTo;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}