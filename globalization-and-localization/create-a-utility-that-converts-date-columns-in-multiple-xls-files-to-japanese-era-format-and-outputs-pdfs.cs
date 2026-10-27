// Title: Batch convert XLS files to PDF while applying Japanese era date format with Aspose.Cells for .NET
// Common Searches: how to format Excel dates as Japanese era when exporting to PDF with Aspose.Cells C# | C# batch process .xls files to PDF and change date cells to Reiwa era format | Aspose.Cells apply custom number format for Japanese era across all worksheets | convert a folder of legacy .xls workbooks to PDF with locale‑specific date style in .NET | set Japanese era date style for DateTime cells before saving workbook as PDF using Aspose.Cells
// Tags: apply Japanese era custom number format Aspose.Cells | batch XLS to PDF conversion Aspose.Cells | date cell styling for Japanese locale .NET | export workbook to PDF with localized dates | process multiple Excel files programmatically C#

using System;
using System.IO;
using Aspose.Cells;
using System.Globalization;

// The utility scans every .xls file in a given input directory, loads each workbook with Aspose.Cells, walks through all worksheets and cells to detect DateTime values, applies a custom number format that renders dates in Japanese era (e.g., "令和3年4月5日"), and then saves each workbook as a PDF in the specified output folder.
class Program
{
    static void Main(string[] args)
    {
        // Input folder containing the XLS files (default "Input" if not provided)
        string inputFolder = args.Length > 0 ? args[0] : "Input";

        // Output folder where the generated PDFs will be saved (default "Output" if not provided)
        string outputFolder = args.Length > 1 ? args[1] : "Output";

        try
        {
            // Ensure the input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder \"{inputFolder}\" does not exist. No files to process.");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Retrieve all .xls files from the input folder (non‑recursive)
            string[] xlsFiles = Directory.GetFiles(inputFolder, "*.xls", SearchOption.TopDirectoryOnly);

            foreach (string xlsPath in xlsFiles)
            {
                try
                {
                    // Verify the file exists before loading
                    if (!File.Exists(xlsPath))
                    {
                        Console.WriteLine($"File not found: {xlsPath}");
                        continue;
                    }

                    // Load the workbook from the XLS file
                    Workbook workbook = new Workbook(xlsPath);

                    // Iterate through each worksheet in the workbook
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        Cells cells = sheet.Cells;
                        int maxRow = cells.MaxDataRow;
                        int maxCol = cells.MaxDataColumn;

                        // Scan every cell within the used range
                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = cells[row, col];

                                // Identify cells that contain a DateTime value
                                if (cell.Type == CellValueType.IsDateTime)
                                {
                                    // Retrieve the current style of the cell
                                    Style style = cell.GetStyle();

                                    // Apply a custom number format that displays the date in Japanese era format.
                                    // Example output: "令和3年4月5日"
                                    style.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"";

                                    // Assign the modified style back to the cell
                                    cell.SetStyle(style);
                                }
                            }
                        }
                    }

                    // Construct the PDF file name based on the original XLS file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(xlsPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as a PDF document
                    workbook.Save(pdfPath, SaveFormat.Pdf);
                    Console.WriteLine($"Converted \"{xlsPath}\" to \"{pdfPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file \"{xlsPath}\": {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
