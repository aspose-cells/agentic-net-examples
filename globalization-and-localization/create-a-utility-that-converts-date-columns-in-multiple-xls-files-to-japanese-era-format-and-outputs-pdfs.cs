// Title: Batch convert XLS workbooks to PDF while converting all date cells to Japanese era format using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans a directory for *.xls files, loads each workbook with Aspose.Cells, sets the workbook culture to ja-JP, formats every DateTime cell using the pattern ggge年M月d日, and saves the workbook as a PDF in a target folder. | Generate a .NET utility that iterates through all worksheets and cells of multiple Excel files, converts date values to Japanese era notation, and exports each file to PDF with Aspose.Cells, ensuring folders exist and handling errors.
// Common Searches: how to batch convert xls files to pdf with Japanese era dates using Aspose.Cells .NET | apply Japanese era custom number format to all date cells in Excel workbook C# | set workbook culture to ja-JP for PDF export with Aspose.Cells | convert multiple Excel files to PDF while changing date format to era in C# | Aspose.Cells example for formatting dates as ggge年M月d日 and saving as PDF
// Tags: Aspose.Cells batch XLS to PDF conversion | Japanese era date formatting Aspose.Cells | ggge年M月d日 custom number format | Workbook CultureInfo ja-JP setting | Iterate worksheets cells for date conversion C#

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

namespace DateToJapaneseEraPdf
{
    // The example scans a specified input folder for .xls files, loads each workbook with Aspose.Cells, sets the workbook culture to Japanese (ja-JP), applies the era pattern ggge年M月d日 to every cell containing a DateTime value, and saves the modified workbook as a PDF in an output folder, with basic error handling for missing files and folder creation.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source XLS files
            string inputFolder = @"C:\Input";

            // Folder where the resulting PDFs will be saved
            string outputFolder = @"C:\Output";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure output directory exists
            Directory.CreateDirectory(outputFolder);

            // Process each XLS file in the input folder
            foreach (string xlsPath in Directory.GetFiles(inputFolder, "*.xls"))
            {
                // Guard against missing file (should not happen with GetFiles)
                if (!File.Exists(xlsPath))
                {
                    Console.WriteLine($"File not found, skipping: {xlsPath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(xlsPath);

                    // Set culture to Japanese to enable era formatting
                    workbook.Settings.CultureInfo = new CultureInfo("ja-JP");

                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through all used cells in the worksheet
                        foreach (Cell cell in sheet.Cells)
                        {
                            // Check if the cell contains a DateTime value
                            if (cell.Value is DateTime)
                            {
                                // Retrieve the cell's style
                                Style style = cell.GetStyle();

                                // Apply Japanese era custom number format (e.g., "ggge年M月d日")
                                style.Custom = "ggge\"年\"M\"月\"d\"日\"";

                                // Assign the modified style back to the cell
                                cell.SetStyle(style);
                            }
                        }
                    }

                    // Build the output PDF file path
                    string pdfFileName = Path.GetFileNameWithoutExtension(xlsPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, SaveFormat.Pdf);

                    Console.WriteLine($"Converted '{xlsPath}' to PDF successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{xlsPath}': {ex.Message}");
                }
            }
        }
    }
}
