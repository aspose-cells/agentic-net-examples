// Title: C# Batch Convert Excel Date Columns to Japanese Era and Export PDFs using Aspose.Cells
// Description: A console utility that scans a folder for .xls/.xlsx workbooks, sets the workbook region to Japan, applies a custom Japanese‑era number format to any column whose first non‑blank cell is a date, saves a temporary XLSX, converts it to PDF with Aspose.Cells ConversionUtility, and cleans up the temporary file.
// Keywords: Aspose.Cells | C# | .NET | Japanese era date format | Excel to PDF batch conversion | custom number format | region Japan | date column styling | automation | ConversionUtility
// Common Searches: format Excel dates as Japanese era using Aspose.Cells | batch convert XLSX to PDF C# Aspose | apply custom number format to whole column Aspose.Cells | set workbook region to Japan for date formatting | convert multiple Excel files to PDF programmatically
// Developer Intent: Process a collection of Excel files, convert any date columns to the Japanese era notation, and generate corresponding PDF documents automatically.
// Use Cases: Modernize legacy Japanese financial spreadsheets by delivering PDFs with era‑based dates for regulatory compliance. | Run a nightly job that formats date columns in newly generated reports and publishes PDF versions for archiving. | Integrate into a CI/CD pipeline to guarantee that all exported PDFs from Excel sources use the correct Japanese calendar representation.
// AI Prompts: Generate C# code that detects date columns in an Excel workbook and applies a Japanese‑era custom format with Aspose.Cells. | Refactor the batch utility to use async I/O and parallel processing for handling thousands of files efficiently. | Explain how to modify the custom format string to show the era name in English or to support other locale‑specific date patterns.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace DateColumnJapaneseEraConverter
{
    // A console utility that scans a folder for .xls/.xlsx workbooks, sets the workbook region to Japan, applies a custom Japanese‑era number format to any column whose first non‑blank cell is a date, saves a temporary XLSX, converts it to PDF with Aspose.Cells ConversionUtility, and cleans up the temporary file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing Excel files (XLS/XLSX)
            string inputFolder = @"C:\InputExcelFiles";
            // Output folder where PDFs will be saved
            string outputFolder = @"C:\OutputPdfFiles";

            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            ProcessFolder(inputFolder, outputFolder);
        }

        static void ProcessFolder(string inputFolder, string outputFolder)
        {
            // Get all Excel files (both .xls and .xlsx) in the folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.xls*");

            foreach (string excelPath in excelFiles)
            {
                try
                {
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(excelPath);

                    // Set workbook region to Japan to ensure Japanese calendar is used
                    workbook.Settings.Region = CountryCode.Japan;

                    // Define the Japanese era custom number format
                    // Example format: "ggge年M月d日" (e.g., "令和3年5月12日")
                    string japaneseEraFormat = "[$-ja-JP]ggge年M月d日";

                    // Create a style with the custom format
                    Style eraStyle = workbook.CreateStyle();
                    eraStyle.Custom = japaneseEraFormat;

                    // Apply the style to each column that contains date values
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Determine the maximum column used in the sheet
                        int maxColumn = sheet.Cells.MaxColumn;

                        for (int col = 0; col <= maxColumn; col++)
                        {
                            // Find the first non‑blank cell in the column
                            int firstDataRow = sheet.Cells.GetFirstDataRow(col);
                            if (firstDataRow < 0) continue; // Column is empty

                            // Check if the first data cell is a date
                            Cell firstCell = sheet.Cells[firstDataRow, col];
                            if (firstCell.Type == CellValueType.IsDateTime)
                            {
                                // Apply the Japanese era style to the whole column
                                StyleFlag flag = new StyleFlag();
                                flag.NumberFormat = true; // enable custom number format
                                sheet.Cells.ApplyColumnStyle(col, eraStyle, flag);
                            }
                        }
                    }

                    // Save the modified workbook to a temporary XLSX file
                    string tempXlsxPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + "_era.xlsx");
                    workbook.Save(tempXlsxPath, SaveFormat.Xlsx);

                    // Convert the temporary XLSX file to PDF using ConversionUtility
                    string pdfPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + ".pdf");
                    ConversionUtility.Convert(tempXlsxPath, pdfPath);

                    // Optionally delete the temporary XLSX file
                    if (File.Exists(tempXlsxPath))
                        File.Delete(tempXlsxPath);

                    Console.WriteLine($"Processed '{excelPath}' -> '{pdfPath}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
    }
}
