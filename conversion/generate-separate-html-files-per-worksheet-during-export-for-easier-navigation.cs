// Title: Export Aspose.Cells workbook to separate HTML files per worksheet using a custom IFilePathProvider (C#)
// AI Prompts: Implement a class that inherits IFilePathProvider and returns a .html filename derived from the worksheet name, then assign it to HtmlSaveOptions.FilePathProvider. | Configure HtmlSaveOptions to export all worksheets, enable ExportWorksheetCSSSeparately, and save the workbook so that an index HTML and individual sheet HTML files are generated. | Write code that creates a workbook with multiple sheets and uses the custom file‑path provider to produce separate HTML pages for each sheet.
// Common Searches: how to export each worksheet as its own HTML file with Aspose.Cells C# | Aspose.Cells C# generate separate HTML pages for each sheet using IFilePathProvider | save workbook to multiple HTML files per sheet Aspose.Cells .NET | separate CSS per worksheet when exporting to HTML Aspose.Cells
// Tags: custom IFilePathProvider HTML export | HtmlSaveOptions per‑worksheet files | Aspose.Cells separate worksheet HTML pages | export Excel to multiple HTML files .NET | worksheet CSS separation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSeparateHtmlExport
{
    // Custom file path provider that creates a separate HTML file for each worksheet
    // The example creates a workbook with three worksheets, defines a CustomFilePathProvider that returns a .html filename based on each sheet's name, configures HtmlSaveOptions to use this provider and to export CSS separately for each sheet, and saves the workbook. The result is an index file (Workbook.html) and individual HTML files (Summary.html, Data.html, Report.html) for each worksheet.
    public class CustomFilePathProvider : IFilePathProvider
    {
        // Returns the file name (including .html extension) based on the worksheet name
        public string GetFullName(string sheetName)
        {
            // You can customize the folder or naming convention here
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample worksheets
            Workbook workbook = new Workbook();

            // First worksheet (default)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";
            sheet1.Cells["A1"].PutValue("This is the summary sheet.");

            // Second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Data");
            sheet2.Cells["A1"].PutValue("Header");
            sheet2.Cells["A2"].PutValue(123);
            sheet2.Cells["B2"].PutValue(456);

            // Third worksheet
            Worksheet sheet3 = workbook.Worksheets.Add("Report");
            sheet3.Cells["A1"].PutValue("Report Content");
            sheet3.Cells["A2"].PutValue(DateTime.Now);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export all worksheets (default), but we explicitly set it for clarity
                ExportActiveWorksheetOnly = false,
                // Use the custom provider to generate a separate file per sheet
                FilePathProvider = new CustomFilePathProvider(),
                // Optional: keep CSS separate for each sheet to reduce file size
                ExportWorksheetCSSSeparately = true
            };

            // Save the workbook. The main file (index) will be created as "Workbook.html"
            // and each worksheet will be saved as "Summary.html", "Data.html", "Report.html"
            workbook.Save("Workbook.html", saveOptions);

            Console.WriteLine("Workbook exported to separate HTML files per worksheet.");
        }
    }
}
