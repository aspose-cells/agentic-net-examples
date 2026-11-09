// Title: Convert multiple Excel workbooks to PDF while applying Japanese era date format using Aspose.Cells for .NET
// Common Searches: how to apply Japanese era (ggge) date format to all date cells in Excel files with Aspose.Cells C# | batch process XLSX files to PDF with custom Japanese era date style in .NET | convert a folder of Excel workbooks to PDF and change date cells to Reiwa era using Aspose.Cells | C# code to iterate through worksheets and set custom number format for DateTime values before PDF export | Aspose.Cells example for localized date formatting during Excel to PDF conversion
// Tags: japanese era custom number format Aspose.Cells | batch convert excel workbooks to pdf .net | apply ggge date format to cells programmatically | iterate worksheets and cells with Aspose.Cells | localized date formatting during pdf export

using System;
using System.IO;
using Aspose.Cells;

namespace DateToJapaneseEraPdf
{
    // The utility scans a specified input folder for .xls and .xlsx files, loads each workbook with Aspose.Cells, walks through every worksheet and cell, applies the custom Japanese era format "ggge年M月d日" to any DateTime value, and then saves the modified workbook as a PDF in an output directory.
    class Program
    {
        // Japanese era custom number format (e.g., Reiwa 3年4月5日)
        private const string JapaneseEraFormat = "ggge年M月d日";

        static void Main(string[] args)
        {
            // Input folder containing XLS/XLSX files
            string inputFolder = @"C:\InputXls";
            // Output folder for generated PDFs
            string outputFolder = @"C:\OutputPdf";

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify input directory exists; if not, inform the user and exit
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input directory not found: {inputFolder}");
                return;
            }

            string[] files;
            try
            {
                // Get all .xls and .xlsx files in the input folder
                files = Directory.GetFiles(inputFolder, "*.xls*", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing input directory: {ex.Message}");
                return;
            }

            foreach (string filePath in files)
            {
                // Verify the file still exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Process each worksheet
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Get the used range of the worksheet
                        Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                        // Iterate through all cells in the used range
                        foreach (Cell cell in usedRange)
                        {
                            // Check if the cell contains a DateTime value
                            if (cell.Value is DateTime)
                            {
                                // Apply Japanese era custom number format
                                Style style = cell.GetStyle();
                                style.Custom = JapaneseEraFormat;
                                cell.SetStyle(style);
                            }
                        }
                    }

                    // Build output PDF file path (same name as source but .pdf extension)
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
                    string pdfPath = Path.Combine(outputFolder, fileNameWithoutExt + ".pdf");

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Converted '{filePath}' to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
    }
}
