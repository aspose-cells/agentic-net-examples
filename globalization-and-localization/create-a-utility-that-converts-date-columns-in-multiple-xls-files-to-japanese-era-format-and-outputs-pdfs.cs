using System;
using System.IO;
using Aspose.Cells;

namespace DateToJapaneseEraPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing XLS files
            string inputFolder = @"C:\InputXls";
            // Output folder for generated PDFs
            string outputFolder = @"C:\OutputPdf";

            // Ensure input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Custom Japanese era date format
            const string japaneseEraFormat = "[$-ja-JP]ggge年M月d日";

            // Process each .xls file in the input folder
            foreach (string xlsPath in Directory.GetFiles(inputFolder, "*.xls"))
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(xlsPath))
                    {
                        Console.WriteLine($"File not found: {xlsPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(xlsPath);

                    // Set regional settings to Japan (helps with era calculations)
                    workbook.Settings.Region = CountryCode.Japan;

                    // Prepare a reusable style with the Japanese era format
                    Style eraStyle = workbook.CreateStyle();
                    eraStyle.Custom = japaneseEraFormat;

                    // Iterate through all worksheets and cells
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        Cells cells = sheet.Cells;
                        int maxRow = cells.MaxDataRow;
                        int maxCol = cells.MaxDataColumn;

                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = cells[row, col];
                                // Apply style to DateTime cells
                                if (cell.Type == CellValueType.IsDateTime)
                                    cell.SetStyle(eraStyle);
                            }
                        }
                    }

                    // Determine the PDF output path
                    string pdfPath = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(xlsPath) + ".pdf");

                    // Save directly to PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);

                    Console.WriteLine($"Converted '{xlsPath}' to PDF at '{pdfPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{xlsPath}': {ex.Message}");
                }
            }
        }
    }
}