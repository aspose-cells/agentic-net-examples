// Title: C# Batch Converter: Excel Dates to Japanese Era & PDF Export with Aspose.Cells
// Description: Scans a folder of XLS/XLSX workbooks, sets the region to Japan, applies a custom Japanese‑era date style to every DateTime cell, and saves each file as a PDF.
// Keywords: Aspose.Cells | Japanese era date format | gengō | Excel to PDF conversion | C# batch processing | globalization | localization Japan | custom date style | region Japan | automated Excel PDF export
// Common Searches: format Excel dates as Japanese era with Aspose.Cells | batch convert Excel files to PDF C# | set workbook region Japan Aspose.Cells | apply custom date style to all cells before PDF export | convert multiple XLSX to PDF with Japanese era dates
// Developer Intent: Automate the conversion of many Excel workbooks so that every date appears in Japanese era notation and each workbook is output as a PDF.
// Use Cases: Generate PDF reports for Japanese clients where dates use the gengō system. | Archive financial spreadsheets with culturally correct era dates before distribution. | Produce localized invoices from Excel templates, converting date columns to era format and exporting them as PDFs.
// AI Prompts: Show how to add detailed error logging for cells that cannot receive the Japanese era style. | Provide a custom format string that includes Japanese weekday names together with the era date. | Explain how to extend the utility to preserve worksheet names as PDF bookmarks while processing .xlsx files.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace DateToJapaneseEraPdfConverter
{
    // Scans a folder of XLS/XLSX workbooks, sets the region to Japan, applies a custom Japanese‑era date style to every DateTime cell, and saves each file as a PDF.
    class Program
    {
        static void Main()
        {
            // Folder containing source Excel files
            string inputFolder = @"C:\InputExcel";
            // Folder where PDF files will be saved
            string outputFolder = @"C:\OutputPdf";

            // Ensure input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the input folder
            foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xls*"))
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Set workbook region to Japan to enable Japanese era formatting
                    workbook.Settings.Region = CountryCode.Japan;

                    // Define a style that formats dates using Japanese era (gengō)
                    Style eraStyle = workbook.CreateStyle();
                    // Custom format: era name (gg), year of era (e), month (M), day (d)
                    eraStyle.Custom = "[$-ja-JP]ggge年M月d日";

                    // Apply the style to all date cells in all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        int maxRow = sheet.Cells.MaxDataRow;
                        int maxCol = sheet.Cells.MaxDataColumn;

                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = sheet.Cells[row, col];
                                if (cell.Type == CellValueType.IsDateTime)
                                {
                                    cell.SetStyle(eraStyle);
                                }
                            }
                        }
                    }

                    // Build the PDF output path
                    string pdfPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + ".pdf");

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Converted: {excelPath} -> {pdfPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Conversion completed.");
        }
    }
}
