// Title: How to export an Aspose.Cells workbook to a UTF‑16 encoded CSV file in C#
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, adds multilingual cells, and saves it as a CSV using UTF‑16 encoding. | Demonstrate configuring TxtSaveOptions for CSV output with Encoding.Unicode in Aspose.Cells. | Show how to verify the output directory exists before writing a UTF‑16 CSV with Aspose.Cells.
// Common Searches: Aspose.Cells C# generate Unicode CSV from Excel workbook | Saving Excel data as UTF‑16 CSV using Aspose.Cells TxtSaveOptions example | Exporting multilingual worksheet to CSV with UTF‑16 encoding in .NET
// Tags: Aspose.Cells TxtSaveOptions CSV UTF-16 | C# export workbook to Unicode CSV | save Excel as UTF-16 CSV Aspose | multilingual CSV export Aspose.Cells | Unicode CSV generation .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // // Demonstrates creating a workbook, inserting sample data including Japanese text, and saving it as a UTF‑16 encoded CSV file using Aspose.Cells TxtSaveOptions.
    public class ExportWorkbookToCsvUtf16
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and add sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Description");
                sheet.Cells["A2"].PutValue("Alice");
                sheet.Cells["B2"].PutValue("こんにちは"); // Japanese greeting to test UTF‑16

                // Create CSV save options and set UTF‑16 (Unicode) encoding
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Encoding = Encoding.Unicode // UTF‑16 encoding
                };

                // Define output file path
                string outputPath = "output_utf16.csv";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as CSV using the specified options
                workbook.Save(outputPath, csvOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorkbookToCsvUtf16.Run();
        }
    }
}
