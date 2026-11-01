// Title: C# example to confirm all Aspose.Cells smart marker tags are resolved after WorkbookDesigner processing
// AI Prompts: Write C# code that loads an Excel workbook, executes WorkbookDesigner.Process, and walks through each worksheet's used cells to list any cells still containing the '<#' and '#>' markers. | Enhance the validation program to automatically create the destination folder if it does not exist before calling workbook.Save. | Adjust the cell‑scanning loop so it skips empty sheets and evaluates only cells of type string for leftover smart‑marker placeholders. | Format the console output to include worksheet name, cell address, and the exact unresolved marker text.
// Common Searches: how to verify that Aspose.Cells smart markers are fully replaced after processing in C# | C# example for finding leftover <# #> tags in an Excel file | Aspose.Cells WorkbookDesigner validation script for .NET developers | method to ensure output directory exists when saving a processed workbook | detect remaining smart marker placeholders after Aspose.Cells processing
// Tags: smart marker resolution validation Aspose.Cells | identify unresolved smart markers C# | WorkbookDesigner processing verification | scan Excel cells for placeholder tags | create output directory before workbook save

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input Excel workbook, runs WorkbookDesigner.Process to replace smart markers, then iterates through each worksheet's used range, checking only string cells for the '<#' and '#>' patterns. It logs any unresolved markers with sheet and cell details, ensures the output folder exists, saves the processed workbook, and includes robust error handling.
class SmartMarkerValidator
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains smart markers
            Workbook workbook = new Workbook(inputPath);

            // Process all smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.Process();

            bool unresolvedFound = false;

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the sheet
                Aspose.Cells.Range? usedRange = cells.MaxDisplayRange;
                if (usedRange == null)
                {
                    continue; // Empty sheet
                }

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Scan each cell within the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Only examine string cells
                        if (cell.Type == CellValueType.IsString)
                        {
                            string value = cell.StringValue;

                            // Unresolved smart markers typically contain the pattern "<#" and "#>"
                            if (!string.IsNullOrEmpty(value) && value.Contains("<#") && value.Contains("#>"))
                            {
                                unresolvedFound = true;
                                Console.WriteLine($"Unresolved smart marker at {sheet.Name}!{cell.Name}: {value}");
                            }
                        }
                    }
                }
            }

            // Report the validation result
            if (!unresolvedFound)
            {
                Console.WriteLine("All smart markers have been resolved.");
            }

            // Ensure output directory exists before saving
            try
            {
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the processed workbook (optional)
                workbook.Save(outputPath);
                Console.WriteLine($"Processed workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
