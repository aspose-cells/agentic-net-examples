// Title: Create lowercase, URL‑friendly worksheet hyperlinks in Aspose.Cells for .NET using a custom IFilePathProvider implementation
// Common Searches: how to generate hyphenated lowercase URLs from Excel worksheet names with Aspose.Cells C# | using IFilePathProvider to add worksheet hyperlinks in Aspose.Cells .NET | Aspose.Cells hyperlink Add method with custom URL provider example | convert worksheet name to web‑safe path for Excel hyperlinks in C# | save workbook with lowercase worksheet links using Aspose.Cells API
// Tags: custom IFilePathProvider Aspose.Cells | lowercase worksheet URL generation C# | hyperlink Add method Aspose.Cells | sanitize worksheet names for web links | Excel workbook save with URL hyperlinks | Aspose.Cells external URL hyperlink

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Simple interface for converting worksheet names to URL paths
    public interface IFilePathProvider
    {
        string GetFilePath(string worksheetName);
    }

    // Implementation that creates a lowercase, URL‑friendly path
    // The example defines an IFilePathProvider interface and a LowercaseFilePathProvider that transforms worksheet names into hyphen‑separated, lower‑case URLs. It creates a workbook, adds worksheets with mixed‑case names, writes a label in cell A1 of each sheet, generates a URL via the provider, and attaches the hyperlink to the cell using Aspose.Cells' Hyperlinks.Add overload before saving the file.
    public class LowercaseFilePathProvider : IFilePathProvider
    {
        public string GetFilePath(string worksheetName)
        {
            if (string.IsNullOrEmpty(worksheetName))
                throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(worksheetName));

            // Replace spaces with hyphens and convert to lower case
            string sanitized = worksheetName.Trim()
                                            .Replace(' ', '-')
                                            .ToLowerInvariant();

            const string baseUrl = "https://example.com/worksheets/";
            return $"{baseUrl}{sanitized}";
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add worksheets with mixed‑case names
                Worksheet ws1 = workbook.Worksheets[0];
                ws1.Name = "SalesData";

                Worksheet ws2 = workbook.Worksheets.Add("Employee Records");
                Worksheet ws3 = workbook.Worksheets.Add("Inventory_List");

                // Instantiate the custom file‑path provider
                IFilePathProvider pathProvider = new LowercaseFilePathProvider();

                // Add a hyperlink in cell A1 of each worksheet
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Ensure the cell exists and set display text
                    ws.Cells["A1"].PutValue("Link to this sheet");

                    // Build the URL using the provider
                    string url = pathProvider.GetFilePath(ws.Name);

                    // Add the hyperlink (type: URL) to cell A1 using the 5‑argument overload
                    // Parameters: firstRow, firstColumn, totalRows, totalColumns, url
                    ws.Hyperlinks.Add(0, 0, 1, 1, url);
                }

                // Determine output path and ensure directory exists
                string outputPath = Path.Combine(Environment.CurrentDirectory, "LowercaseWorksheetLinks.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to: {outputPath}");
                Console.WriteLine("Hyperlinks have been created using lowercase URLs derived from worksheet names.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
