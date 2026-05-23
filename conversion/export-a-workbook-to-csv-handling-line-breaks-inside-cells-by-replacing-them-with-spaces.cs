using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportCsvWithLineBreakHandling
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data that contains line breaks
                sheet.Cells["A1"].PutValue("First line\nSecond line");
                sheet.Cells["B2"].PutValue("Row1\r\nRow2\rRow3");
                sheet.Cells["C3"].PutValue("NoBreaksHere");

                // Replace line break characters in all string cells with a space
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            if (!string.IsNullOrEmpty(original))
                            {
                                string cleaned = original.Replace("\r\n", " ")
                                                         .Replace("\n", " ")
                                                         .Replace("\r", " ");
                                // Update the cell only if changes were made
                                if (cleaned != original)
                                {
                                    cell.PutValue(cleaned);
                                }
                            }
                        }
                    }
                }

                // Configure CSV save options
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    Separator = ',',               // Use comma as delimiter
                    KeepSeparatorsForBlankRow = false,
                    ExportAllSheets = false        // Export only the active sheet
                };

                // Ensure the output directory exists
                string outputPath = "output.csv";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to CSV; line breaks are now spaces
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"CSV file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportCsvWithLineBreakHandling.Run();
        }
    }
}