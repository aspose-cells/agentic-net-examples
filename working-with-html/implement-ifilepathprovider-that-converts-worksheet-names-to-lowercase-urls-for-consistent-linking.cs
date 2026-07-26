// Title: Custom IFilePathProvider for Lower‑Case Worksheet HTML Files in Aspose.Cells (.NET)
// Description: Demonstrates how to implement a custom IFilePathProvider that sanitizes worksheet names with CellsHelper.CreateSafeSheetName, converts them to lower‑case, adds a .html extension, and assigns the provider to HtmlSaveOptions so each sheet is saved as a predictable, lower‑case HTML file.
// Keywords: Aspose.Cells IFilePathProvider | lowercase HTML filenames | worksheet to HTML export | CreateSafeSheetName | custom file path provider .NET | HTMLSaveOptions example | case‑insensitive URLs Excel
// Common Searches: Aspose.Cells custom file path provider example | export each worksheet to lower case html file | convert worksheet names to lowercase URLs Aspose.Cells | how to use IFilePathProvider with HtmlSaveOptions | sanitize sheet names for html export Aspose.Cells
// Developer Intent: Create a reusable IFilePathProvider that generates safe, lower‑case .html file names from worksheet names for consistent linking.
// Use Cases: Generate SEO‑friendly, case‑insensitive URLs for each worksheet when publishing workbooks as HTML. | Handle sheet names containing spaces, capitals, or illegal characters by normalizing them before file creation. | Integrate into web portals that serve individual worksheet pages with predictable file paths.
// AI Prompts: Write a C# class implementing Aspose.Cells IFilePathProvider that returns lower‑case .html file names using CellsHelper.CreateSafeSheetName. | Show how to configure HtmlSaveOptions to use a custom file path provider for exporting each worksheet to a separate lower‑case HTML file. | Explain best practices for sanitizing worksheet names and generating URL‑safe file names during HTML export with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom implementation of IFilePathProvider that converts worksheet names to lowercase URLs.
    // Demonstrates how to implement a custom IFilePathProvider that sanitizes worksheet names with CellsHelper.CreateSafeSheetName, converts them to lower‑case, adds a .html extension, and assigns the provider to HtmlSaveOptions so each sheet is saved as a predictable, lower‑case HTML file.
    internal class LowerCaseFilePathProvider : IFilePathProvider
    {
        // Returns a file name based on the worksheet name, converted to lower case and suffixed with .html.
        // Uses CellsHelper.CreateSafeSheetName to ensure the name is a valid Excel sheet name before conversion.
        public string GetFullName(string sheetName)
        {
            // Ensure the sheet name is safe for file naming.
            string safeName = CellsHelper.CreateSafeSheetName(sheetName);
            // Convert to lower case and add the .html extension.
            return $"{safeName.ToLowerInvariant()}.html";
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add sample data.
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "SalesData";
            sheet1.Cells["A1"].PutValue("Quarter");
            sheet1.Cells["B1"].PutValue("Revenue");
            sheet1.Cells["A2"].PutValue("Q1");
            sheet1.Cells["B2"].PutValue(15000);

            // Add a second worksheet with a name that contains spaces and uppercase letters.
            Worksheet sheet2 = workbook.Worksheets.Add("Employee Summary");
            sheet2.Cells["A1"].PutValue("Name");
            sheet2.Cells["B1"].PutValue("Department");
            sheet2.Cells["A2"].PutValue("Alice");
            sheet2.Cells["B2"].PutValue("Finance");

            // Configure HTML save options to use the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export each worksheet to a separate HTML file.
                ExportActiveWorksheetOnly = false,
                // Use the custom provider that generates lower‑case URLs.
                FilePathProvider = new LowerCaseFilePathProvider()
            };

            // Save the workbook; each worksheet will be saved as a lower‑case HTML file.
            workbook.Save("WorkbookOutput.html", saveOptions);

            Console.WriteLine("Workbook saved with lower‑case HTML file names.");
        }
    }
}
