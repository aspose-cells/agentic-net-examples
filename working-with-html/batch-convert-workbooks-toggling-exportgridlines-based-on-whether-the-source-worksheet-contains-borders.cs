using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

class BatchConvertWorkbooks
{
    static void Main()
    {
        // Folder containing source Excel files
        string sourceFolder = @"C:\Input";
        // Folder where converted HTML files will be saved
        string targetFolder = @"C:\Output";

        try
        {
            // Ensure the output directory exists
            Directory.CreateDirectory(targetFolder);

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Process each .xlsx file in the source folder
            foreach (string sourcePath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Skip if the file somehow does not exist
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"File not found, skipping: {sourcePath}");
                    continue;
                }

                try
                {
                    // Load the workbook (create & load lifecycle)
                    Workbook workbook = new Workbook(sourcePath);

                    bool hasBorder = false;

                    // Scan worksheets for any cell that has a border applied
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        // Determine the used range limits
                        int maxRow = ws.Cells.MaxDataRow;
                        int maxCol = ws.Cells.MaxDataColumn;

                        for (int row = 0; row <= maxRow && !hasBorder; row++)
                        {
                            for (int col = 0; col <= maxCol && !hasBorder; col++)
                            {
                                Style style = ws.Cells[row, col].GetStyle();

                                // Check all possible border sides
                                if (style.Borders[BorderType.TopBorder].LineStyle != CellBorderType.None ||
                                    style.Borders[BorderType.BottomBorder].LineStyle != CellBorderType.None ||
                                    style.Borders[BorderType.LeftBorder].LineStyle != CellBorderType.None ||
                                    style.Borders[BorderType.RightBorder].LineStyle != CellBorderType.None ||
                                    style.Borders[BorderType.DiagonalDown].LineStyle != CellBorderType.None ||
                                    style.Borders[BorderType.DiagonalUp].LineStyle != CellBorderType.None)
                                {
                                    hasBorder = true;
                                }
                            }
                        }

                        if (hasBorder) break;
                    }

                    // Configure HTML save options; ExportGridLines follows border presence
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        ExportGridLines = hasBorder,
                        ExportActiveWorksheetOnly = true
                    };

                    // Destination HTML file path
                    string destPath = Path.Combine(targetFolder,
                        Path.GetFileNameWithoutExtension(sourcePath) + ".html");

                    // Convert using ConversionUtility (leverages provided conversion rule)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                    ConversionUtility.Convert(sourcePath, loadOptions, destPath, htmlOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}