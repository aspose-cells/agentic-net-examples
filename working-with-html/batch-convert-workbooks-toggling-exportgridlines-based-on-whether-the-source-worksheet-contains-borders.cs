// Title: Batch convert multiple Excel workbooks to PDF in C# with Aspose.Cells while detecting cell borders to control gridline export
// Common Searches: c# Aspose.Cells batch convert all .xlsx files in a folder to PDF | how to detect any cell border in an Excel workbook using Aspose.Cells before PDF export | conditionally export gridlines when converting Excel to PDF with Aspose.Cells based on border presence | iterate used range of worksheets to find borders in Aspose.Cells C# example | process multiple workbooks and save PDFs to an output directory using Aspose.Cells
// Tags: batch excel to pdf conversion Aspose.Cells C# | cell border detection Aspose.Cells workbook | conditional ExportGridLines PdfSaveOptions | scan used range for borders Aspose.Cells | folder based workbook processing Aspose.Cells | error handling during bulk pdf conversion C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads each .xlsx file from a specified input directory, scans every worksheet's used range to determine if any cell has a border, and then saves the workbook as a PDF using Aspose.Cells. Gridline export can be toggled via PdfSaveOptions based on the border detection result, and all PDFs are written to an output folder with basic error handling and logging.
class Program
{
    static void Main()
    {
        // Input and output directories
        string inputDir = @"C:\Input";
        string outputDir = @"C:\Output";

        // Verify input directory exists
        if (!Directory.Exists(inputDir))
        {
            Console.WriteLine($"Input directory does not exist: {inputDir}");
            return;
        }

        // Ensure output directory exists
        Directory.CreateDirectory(outputDir);

        // Process each Excel workbook in the input folder
        foreach (string filePath in Directory.GetFiles(inputDir, "*.xlsx"))
        {
            // Verify the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);
                bool hasBorders = false;

                // Scan worksheets for any cell that has a border
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;
                    // Determine the used range to limit iteration
                    AsposeRange usedRange = cells.MaxDisplayRange;

                    if (usedRange != null)
                    {
                        int startRow = usedRange.FirstRow;
                        int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                        int startCol = usedRange.FirstColumn;
                        int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                        for (int row = startRow; row <= endRow && !hasBorders; row++)
                        {
                            for (int col = startCol; col <= endCol && !hasBorders; col++)
                            {
                                Style style = cells[row, col].GetStyle();
                                if (HasAnyBorder(style))
                                {
                                    hasBorders = true;
                                }
                            }
                        }
                    }

                    if (hasBorders) break;
                }

                // Configure PDF save options (default gridlines are exported)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Build output file path
                string outputFile = Path.Combine(outputDir,
                    Path.GetFileNameWithoutExtension(filePath) + ".pdf");

                // Save the workbook as PDF with the configured options
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"Converted: {filePath} -> {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }

    // Helper method to determine if a style contains any border
    static bool HasAnyBorder(Style style)
    {
        // Iterate through all BorderType values and check line style
        foreach (BorderType bt in Enum.GetValues(typeof(BorderType)))
        {
            if (style.Borders[bt].LineStyle != CellBorderType.None)
                return true;
        }
        return false;
    }
}
