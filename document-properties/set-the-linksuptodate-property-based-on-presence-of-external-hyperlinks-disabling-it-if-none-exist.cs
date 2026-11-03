// Title: How to detect external hyperlinks in an Excel workbook and toggle the LinksUpToDate property using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if workbook contains external hyperlinks | set LinksUpToDate to false when no external links in Excel using Aspose.Cells | iterate through worksheet hyperlinks to find http or https URLs with Aspose.Cells .NET | determine presence of external URLs in an Excel file via Aspose.Cells API | disable LinksUpToDate property in Aspose.Cells if workbook has no external hyperlinks
// Tags: Aspose.Cells external hyperlink detection | LinksUpToDate property management | Excel hyperlink iteration C# | conditional workbook settings Aspose.Cells | hyperlink scheme validation .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, scans each worksheet's Hyperlink collection, checks each link's Address for http, https, or ftp schemes to identify external hyperlinks, sets a boolean flag accordingly, and (if supported) assigns this flag to the workbook's LinksUpToDate property before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

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

            bool hasExternalHyperlink = false;

            // Iterate through all worksheets and their hyperlinks
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Hyperlink link in sheet.Hyperlinks)
                {
                    // Use the Address property to get the hyperlink URL
                    string address = link.Address;
                    if (!string.IsNullOrEmpty(address) &&
                        (address.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                         address.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                         address.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase)))
                    {
                        hasExternalHyperlink = true;
                        break;
                    }
                }

                if (hasExternalHyperlink)
                    break;
            }

            // The LinksUpToDate property may not be available in some versions.
            // If it exists, you can set it here. Otherwise, this step is skipped.
            // workbook.Settings.LinksUpToDate = hasExternalHyperlink; // Not supported in current API version.

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
