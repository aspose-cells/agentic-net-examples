// Title: Programmatically replace Excel hyperlink UNC paths with a new network share using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx workbook with Aspose.Cells, iterate through each worksheet's HyperlinkCollection, and change any hyperlink address that begins with \\oldserver\share\ to the same relative path prefixed by \\newserver\share\. | After updating all external link URLs, save the workbook to a new file location, ensuring the output directory exists and handling missing input files gracefully.
// Common Searches: C# Aspose.Cells update all worksheet hyperlinks to a new UNC share path | How to change network share prefix in Excel hyperlinks using Aspose.Cells .NET | Replace old server path with new server in Excel workbook hyperlinks programmatically | Batch modify external link URLs in an .xlsx file with Aspose.Cells | Aspose.Cells iterate HyperlinkCollection and edit Address property in C#
// Tags: Aspose.Cells replace hyperlink UNC path | C# update Excel hyperlink addresses | HyperlinkCollection address modification Aspose.Cells | batch hyperlink migration .xlsx | programmatic external link update Excel .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, walks through every worksheet's HyperlinkCollection, and substitutes any hyperlink that starts with the old UNC prefix (\\oldserver\share\) with the new prefix (\\newserver\share\) while preserving the remainder of the path. It then saves the modified workbook to a new file, handling missing input files and creating the output directory if needed.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = @"C:\Input\MyWorkbook.xlsx";
        string outputPath = @"C:\Output\MyWorkbook_Updated.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Ensure the output directory exists
        try
        {
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to prepare output directory: {ex.Message}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the old and new network share prefixes
            const string oldPrefix = @"\\oldserver\share\";
            const string newPrefix = @"\\newserver\share\";

            // Iterate through all worksheets and update hyperlink addresses
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Hyperlink collection for the current worksheet
                HyperlinkCollection hyperlinks = sheet.Hyperlinks;

                foreach (Hyperlink link in hyperlinks)
                {
                    // Get the current hyperlink address (URL or file path)
                    string? address = link.Address;

                    // If the address starts with the old network share, replace it
                    if (!string.IsNullOrEmpty(address) &&
                        address.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        string updatedAddress = newPrefix + address.Substring(oldPrefix.Length);
                        link.Address = updatedAddress;
                    }
                }
            }

            // Save the workbook with updated hyperlink URLs
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
