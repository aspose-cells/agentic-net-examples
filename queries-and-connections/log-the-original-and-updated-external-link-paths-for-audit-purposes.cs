// Title: Audit and update external link URLs in an Excel workbook using Aspose.Cells for .NET with reflection
// Common Searches: how to audit external link URLs in an Excel file with Aspose.Cells C# | using reflection to change external link paths in an Aspose.Cells workbook | log original and updated external links when saving Excel workbook in .NET | replace old folder path in external links of Excel using Aspose.Cells | save workbook after modifying external link URLs programmatically C# | extract and rewrite external link URLs from worksheets with Aspose.Cells
// Tags: external link URL audit Aspose.Cells | reflection based external link path update | log external link modifications C# | programmatic worksheet external link editing | save workbook after external URL change | replace base directory in external link URLs

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an existing Excel workbook, uses reflection to safely access each worksheet's ExternalLinks collection, reads each link's Url, applies a custom rule to replace an old base folder with a new one, updates the Url when it differs, records the worksheet name along with original and updated paths in an audit log, saves the modified workbook, and writes the audit information to a text file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";
        const string auditFile = "ExternalLinksAudit.txt";

        // Verify that the input workbook exists
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Prepare a list to hold audit information
            List<string> auditLog = new List<string>
            {
                "External Links Audit Log",
                $"Timestamp: {DateTime.Now}"
            };

            // Iterate through all worksheets and process external links if the API is available
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // The ExternalLinks collection may not be present in older Aspose.Cells versions.
                // Guard against missing members using reflection.
                var externalLinksProp = typeof(Worksheet).GetProperty("ExternalLinks");
                if (externalLinksProp == null) continue;

                var linksObj = externalLinksProp.GetValue(sheet);
                if (linksObj == null) continue;

                // The collection is expected to be of type ExternalLinkCollection
                var linkCollection = linksObj as IEnumerable<object>;
                if (linkCollection == null) continue;

                foreach (var linkItem in linkCollection)
                {
                    // Use reflection to access the Url property safely
                    var urlProp = linkItem.GetType().GetProperty("Url");
                    if (urlProp == null) continue;

                    string originalPath = urlProp.GetValue(linkItem) as string ?? string.Empty;
                    string updatedPath = GetUpdatedPath(originalPath);

                    // Update the Url if it differs
                    if (!originalPath.Equals(updatedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        urlProp.SetValue(linkItem, updatedPath);
                    }

                    // Log the changes
                    auditLog.Add($"Worksheet: {sheet.Name}");
                    auditLog.Add($"Original: {originalPath}");
                    auditLog.Add($"Updated : {updatedPath}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputFile, SaveFormat.Xlsx);

            // Write the audit log to a text file
            File.WriteAllLines(auditFile, auditLog);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method that defines how an external link path should be updated
    static string GetUpdatedPath(string originalPath)
    {
        // Example rule: replace an old base folder with a new one
        const string oldBase = @"C:\OldData";
        const string newBase = @"D:\NewData";

        if (originalPath.StartsWith(oldBase, StringComparison.OrdinalIgnoreCase))
        {
            return newBase + originalPath.Substring(oldBase.Length);
        }

        // No change needed
        return originalPath;
    }
}
