// Title: Aspose.Cells .NET – Custom IFilePathProvider for Unique HTML Filenames When Exporting Worksheets
// Description: This example shows how to implement a custom IFilePathProvider that creates safe, case‑insensitive sheet names and appends a numeric suffix only when a worksheet title repeats. The provider is used with HtmlSaveOptions to export each worksheet to a separate HTML file without filename collisions, even when sheet titles contain illegal file‑system characters.
// Keywords: Aspose.Cells | IFilePathProvider | unique HTML filenames | duplicate worksheet names | HtmlSaveOptions | CreateSafeSheetName | .NET | C# | export workbook to HTML | file name collision handling | case‑insensitive dictionary
// Common Searches: Aspose.Cells custom IFilePathProvider example | generate unique HTML file names for each worksheet | avoid duplicate sheet name collisions when saving to HTML | export workbook to multiple HTML files Aspose.Cells .NET | CreateSafeSheetName usage in Aspose.Cells
// Developer Intent: Create a custom IFilePathProvider that returns a distinct HTML file name for every worksheet, handling duplicate or unsafe sheet titles.
// Use Cases: Export a workbook with several sheets named "Report" to separate HTML files without overwriting any file. | Integrate the provider into a web service that generates per‑sheet HTML reports, guaranteeing unique filenames on the server. | Sanitize sheet titles containing characters illegal for file systems while still producing unique filenames for each exported HTML page.
// AI Prompts: Write a C# class that implements IFilePathProvider, adds numeric suffixes to duplicate sheet names, and works with HtmlSaveOptions to export each worksheet as a separate HTML file. | Show how to configure HtmlSaveOptions to use a custom IFilePathProvider for unique HTML filenames when saving a workbook with Aspose.Cells. | Explain the role of CellsHelper.CreateSafeSheetName in preventing invalid characters in generated HTML filenames.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom implementation of IFilePathProvider that ensures unique file names
    // even when worksheets have duplicate (case‑insensitive) titles.
    // This example shows how to implement a custom IFilePathProvider that creates safe, case‑insensitive sheet names and appends a numeric suffix only when a worksheet title repeats. The provider is used with HtmlSaveOptions to export each worksheet to a separate HTML file without filename collisions, even when sheet titles contain illegal file‑system characters.
    internal class UniqueFilePathProvider : IFilePathProvider
    {
        // Tracks how many times a particular safe sheet name has been used.
        private readonly Dictionary<string, int> _nameUsage = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public string GetFullName(string sheetName)
        {
            // Convert the original sheet name to a safe file‑system name.
            string safeName = CellsHelper.CreateSafeSheetName(sheetName);

            // Determine the current usage count for this safe name.
            if (_nameUsage.TryGetValue(safeName, out int count))
            {
                // Increment the count and store it back.
                count++;
                _nameUsage[safeName] = count;
            }
            else
            {
                // First occurrence of this name.
                count = 0;
                _nameUsage[safeName] = count;
            }

            // Append a numeric suffix only when a duplicate exists (count > 0).
            string fileName = count == 0 ? $"{safeName}.html" : $"{safeName}_{count}.html";

            // Return the full file name (relative path). Adjust as needed for absolute paths.
            return fileName;
        }
    }

    public class IFilePathProviderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a workbook.
                Workbook workbook = new Workbook();

                // First sheet (default name "Sheet1").
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Report";

                // Add a second sheet with a distinct internal name.
                Worksheet sheet2 = workbook.Worksheets.Add("Report_Second");

                // Add a third sheet with another distinct name.
                Worksheet sheet3 = workbook.Worksheets.Add("Report_Third");

                // Populate some data (optional).
                sheet1.Cells["A1"].PutValue("Data in first Report sheet");
                sheet2.Cells["A1"].PutValue("Data in second Report sheet");
                sheet3.Cells["A1"].PutValue("Data in third Report sheet");

                // Configure HTML save options to use the custom file‑path provider.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    // Export each worksheet to a separate HTML file.
                    ExportActiveWorksheetOnly = false,
                    // Assign the custom provider.
                    FilePathProvider = new UniqueFilePathProvider()
                };

                // Save the workbook. The provider will generate unique file names such as:
                // Report.html, Report_Second.html, Report_Third.html
                workbook.Save("CombinedOutput.html", saveOptions);

                Console.WriteLine("Workbook saved with unique HTML file names for each worksheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            IFilePathProviderDemo.Run();
        }
    }
}
