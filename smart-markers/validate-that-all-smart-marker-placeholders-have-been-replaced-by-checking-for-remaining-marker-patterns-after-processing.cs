// Title: C# code to confirm that all Aspose.Cells smart marker placeholders are replaced after WorkbookDesigner processing
// Common Searches: how to check if smart markers were fully replaced in an Excel file using Aspose.Cells C# | C# example for validating no remaining &= placeholders after WorkbookDesigner.Process | detect unreplaced smart marker patterns in Aspose.Cells worksheets with regex | verify smart marker replacement in a processed workbook using Aspose.Cells API | automated validation of smart marker placeholders after processing in C#
// Tags: smart marker replacement validation Aspose.Cells | WorkbookDesigner processing verification C# | regex detection of Aspose smart marker placeholders | iterate worksheet cells for placeholder check | unreplaced smart markers detection Excel

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an Excel workbook, runs WorkbookDesigner.Process to replace smart markers, then scans every used cell with a compiled regex ("&=\w+") to locate any remaining placeholders, reporting their sheet, cell address, and whether validation succeeded.
class SmartMarkerValidator
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains smart markers
            string inputPath = "InputWithSmartMarkers.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.Process();

            // Regex pattern that matches typical smart marker placeholders (e.g., "&=CustomerName")
            Regex markerPattern = new Regex(@"&=\w+", RegexOptions.Compiled);

            bool markersFound = false;

            // Iterate through all worksheets and cells to detect any remaining markers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the sheet to limit the search
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // If the sheet is empty, skip it
                if (usedRange == null)
                    continue;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Only check cells that contain string values
                        if (cell.Type == CellValueType.IsString)
                        {
                            string cellText = cell.StringValue;
                            if (markerPattern.IsMatch(cellText))
                            {
                                markersFound = true;
                                Console.WriteLine(
                                    $"Unreplaced smart marker found in sheet '{sheet.Name}' at cell {cell.Name}: \"{cellText}\"");
                            }
                        }
                    }
                }
            }

            if (!markersFound)
            {
                Console.WriteLine("Validation passed: No smart marker placeholders remain after processing.");
            }
            else
            {
                Console.WriteLine("Validation failed: Some smart marker placeholders were not replaced.");
            }

            // Optionally, save the processed workbook (if needed)
            // string outputPath = "ProcessedOutput.xlsx";
            // workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
