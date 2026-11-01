// Title: Generate a tab‑separated audit log of each non‑empty cell’s address, type, and value while iterating worksheets with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates over each worksheet’s populated area, and records every non‑empty cell’s address, data type, and value into a tab‑separated text file. | Enhance the enumeration to prepend the worksheet name to each logged line and format any DateTime values using ISO‑8601 before writing them. | Add comprehensive exception handling that captures errors occurring during cell processing and appends the details to a separate error‑log file while allowing the audit to continue.
// Common Searches: how to export cell address, type and value from Excel using Aspose.Cells C# | Aspose.Cells iterate used range and write audit report to txt file | skip empty cells when creating Excel audit report with Aspose.Cells .NET | log worksheet name and cell metadata to tab separated file using Aspose.Cells | C# Aspose.Cells create audit trail of workbook contents
// Tags: Aspose.Cells enumerate used range | write cell metadata to tab delimited file | audit non empty Excel cells C# | skip null cells Aspose.Cells | worksheet level cell address logging

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads an Excel workbook with Aspose.Cells, determines the populated range of each worksheet, iterates through all non‑empty cells, and writes each cell's address, data type, and stringified value to a tab‑separated audit.txt file, creating the output directory if necessary and handling missing files gracefully.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook to be audited
            string workbookPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(workbookPath)}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path for the audit log file
            string logPath = "audit.txt";

            // Ensure the directory for the log file exists
            string? logDir = Path.GetDirectoryName(Path.GetFullPath(logPath));
            if (!string.IsNullOrEmpty(logDir) && !Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            // Open a StreamWriter to write the audit information
            using (StreamWriter writer = new StreamWriter(logPath, false))
            {
                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    writer.WriteLine($"Worksheet: {sheet.Name}");

                    // Determine the used range to limit enumeration to populated cells
                    AsposeRange? usedRange = sheet.Cells.MaxDisplayRange;

                    // If the sheet is empty, skip to next worksheet
                    if (usedRange == null)
                    {
                        writer.WriteLine(); // Separate worksheets with a blank line
                        continue;
                    }

                    // Calculate start and end indices from the Range object
                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Loop through rows and columns within the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Skip cells that have no value
                            if (cell.Type == CellValueType.IsNull)
                                continue;

                            // Cell address (e.g., A1)
                            string address = cell.Name;

                            // Data type of the cell
                            string dataType = cell.Type.ToString();

                            // Cell value as string (handle nulls)
                            string value = cell.Value?.ToString() ?? string.Empty;

                            // Write the audit entry: address, data type, value
                            writer.WriteLine($"{address}\t{dataType}\t{value}");
                        }
                    }

                    writer.WriteLine(); // Separate worksheets with a blank line
                }
            }

            Console.WriteLine($"Audit log created at: {Path.GetFullPath(logPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
