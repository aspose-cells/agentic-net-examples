// Title: C# example to confirm all Aspose.Cells smart markers are resolved after WorkbookDesigner processing
// Common Searches: how to check if smart markers were resolved after calling WorkbookDesigner.Process in C# | detect remaining '&=' smart marker tags in an Excel file using Aspose.Cells | validate smart marker resolution in a workbook after processing with Aspose.Cells | C# code to scan worksheets for unresolved smart marker patterns in Excel
// Tags: smart marker resolution validation Aspose.Cells | WorkbookDesigner processing verification C# | detect unresolved &= markers in Excel | iterate worksheet cells for smart marker check | post‑processing smart marker verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, runs WorkbookDesigner.Process to resolve smart markers, then iterates every cell in each worksheet to look for the '&=' prefix that indicates an unresolved smart marker, reports the result, and saves the processed file.
class SmartMarkerValidator
{
    static void Main(string[] args)
    {
        try
        {
            // Path to the source workbook containing smart markers
            string sourcePath = "InputWithSmartMarkers.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(sourcePath);

            // Process smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.Process(); // resolves all smart markers

            // Validate that no unresolved smart marker tags remain
            bool allResolved = AreSmartMarkersResolved(workbook);
            Console.WriteLine(allResolved
                ? "All smart marker tags have been resolved."
                : "Unresolved smart marker tags were found.");

            // Optionally save the processed workbook (lifecycle rule: save)
            string outputPath = "ProcessedOutput.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Processed workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Checks every cell in every worksheet for unresolved smart marker patterns.
    private static bool AreSmartMarkersResolved(Workbook workbook)
    {
        // Unresolved smart markers typically start with "&=" (e.g., "&=Customer.Name")
        const string unresolvedPrefix = "&=";

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range; if null, skip the sheet
            Aspose.Cells.Range range = sheet.Cells.MaxDisplayRange;
            if (range == null)
                continue;

            int startRow = range.FirstRow;
            int endRow = range.FirstRow + range.RowCount - 1;
            int startCol = range.FirstColumn;
            int endCol = range.FirstColumn + range.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;
                        if (!string.IsNullOrEmpty(text) && text.Contains(unresolvedPrefix))
                        {
                            // Unresolved smart marker found
                            return false;
                        }
                    }
                }
            }
        }

        // No unresolved markers detected
        return true;
    }
}
