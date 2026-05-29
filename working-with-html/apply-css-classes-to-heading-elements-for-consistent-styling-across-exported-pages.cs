using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ApplyCssToHeadingsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Header 1");
                sheet.Cells["B1"].PutValue("Header 2");
                sheet.Cells["A2"].PutValue("Data 1");
                sheet.Cells["B2"].PutValue("Data 2");

                // Configure HTML save options with custom CSS for headings
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportRowColumnHeadings = true,
                    SaveAsSingleFile = true,
                    CssStyles = @"
.heading {
    font-weight: bold;
    color: #2A7AE2;
    background-color: #E8F0FE;
    text-align: center;
}"
                };

                // Define output path
                string outputPath = "WorkbookWithStyledHeadings.html";

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save workbook as HTML
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"HTML file saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyCssToHeadingsDemo.Run();
        }
    }
}