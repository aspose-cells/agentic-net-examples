using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PdfBookmarkWithDestinationNameDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and add a worksheet named "Report"
                Workbook workbook = new Workbook();
                Worksheet reportSheet = workbook.Worksheets[0];
                reportSheet.Name = "Report";

                // Add some sample content to the worksheet
                reportSheet.Cells["A1"].PutValue("Annual Report Overview");
                reportSheet.Cells["A2"].PutValue("Details follow...");

                // Define output file path
                string outputPath = "ReportWithBookmark.pdf";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF file
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine("PDF created successfully at: " + Path.GetFullPath(outputPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}