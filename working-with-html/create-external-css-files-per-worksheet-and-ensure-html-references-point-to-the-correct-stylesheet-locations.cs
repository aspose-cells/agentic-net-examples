using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom file path provider to control the names of generated HTML and CSS files per worksheet
    public class CustomFilePathProvider : IFilePathProvider
    {
        // Returns the base file name (without extension) for a given worksheet.
        // Aspose.Cells will append appropriate extensions for HTML and CSS files.
        public string GetFullName(string sheetName)
        {
            // Example: use the worksheet name as the file base name, sanitizing any invalid characters.
            string safeName = sheetName.Replace(" ", "_").Replace("/", "_");
            return safeName;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Report");

            // Populate each worksheet with sample data
            workbook.Worksheets["Summary"].Cells["A1"].PutValue("Overview");
            workbook.Worksheets["Data"].Cells["A1"].PutValue("Item");
            workbook.Worksheets["Data"].Cells["B1"].PutValue("Quantity");
            workbook.Worksheets["Report"].Cells["A1"].PutValue("Result");

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export a separate CSS file for each worksheet
                ExportWorksheetCSSSeparately = true,

                // Use the custom file path provider so that each worksheet's HTML and CSS files
                // are named after the worksheet (e.g., Summary.html, Summary.css, etc.)
                FilePathProvider = new CustomFilePathProvider()
            };

            // Define the output directory
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AsposeHtmlExport");
            Directory.CreateDirectory(outputDir);

            // Save the workbook as HTML. Because ExportWorksheetCSSSeparately is true,
            // Aspose.Cells will generate:
            //   Summary.html, Summary.css
            //   Data.html, Data.css
            //   Report.html, Report.css
            // All files will be placed in the output directory.
            string outputPath = Path.Combine(outputDir, "index.html"); // primary index file
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("HTML export completed. Files are located at:");
            Console.WriteLine(outputDir);
        }
    }
}