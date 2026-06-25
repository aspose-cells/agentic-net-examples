using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Utility;

namespace AsposeCellsDemo
{
    public class BatchExcelToPdfWithTimeline
    {
        // Convert zero‑based column index to Excel column letter (e.g., 0 -> "A")
        private static string GetColumnLetter(int index)
        {
            string column = string.Empty;
            while (index >= 0)
            {
                column = (char)('A' + (index % 26)) + column;
                index = index / 26 - 1;
            }
            return column;
        }

        public static void Run(string inputFolder, string outputFolder)
        {
            try
            {
                // Ensure output directory exists
                Directory.CreateDirectory(outputFolder);

                // Process each .xlsx file in the input folder
                foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
                {
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(excelPath);
                        Worksheet sheet = workbook.Worksheets[0];

                        // Determine the used range of the worksheet
                        int lastRow = sheet.Cells.MaxDataRow;      // zero‑based
                        int lastCol = sheet.Cells.MaxDataColumn;   // zero‑based
                        string startCell = "A1";
                        string endCell = $"{GetColumnLetter(lastCol)}{lastRow + 1}";
                        string dataRange = $"{startCell}:{endCell}";

                        // Add a pivot table based on the used range
                        int pivotIndex = sheet.PivotTables.Add(dataRange, "E1", "PivotTable1");
                        PivotTable pivot = sheet.PivotTables[pivotIndex];

                        // Use the first column as the row field (commonly a date field)
                        pivot.AddFieldToArea(PivotFieldType.Row, 0);

                        // Add a timeline control linked to the pivot table.
                        sheet.Timelines.Add(pivot, "G1", 0);

                        // Save the workbook as PDF
                        string pdfPath = Path.Combine(outputFolder,
                            Path.GetFileNameWithoutExtension(excelPath) + ".pdf");
                        workbook.Save(pdfPath, SaveFormat.Pdf);

                        Console.WriteLine($"Converted '{excelPath}' to PDF successfully.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Provide input and output folders via command‑line arguments or hard‑code them.
            string inputFolder = args.Length > 0 ? args[0] : @"C:\InputExcels";
            string outputFolder = args.Length > 1 ? args[1] : @"C:\OutputPdfs";

            BatchExcelToPdfWithTimeline.Run(inputFolder, outputFolder);
        }
    }
}