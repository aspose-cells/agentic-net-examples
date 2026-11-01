// Title: Hide numeric zero values in all worksheets of multiple Excel workbooks from a folder using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells batch hide zero cells in every worksheet of a folder of .xlsx files | How to clear numeric zeros across all sheets when processing multiple Excel workbooks with Aspose.Cells | Programmatically remove zero values from Excel files in a directory using Aspose.Cells .NET | Iterate through worksheets and cells to delete zero entries in bulk with Aspose.Cells C# | Skip password‑protected Excel files while batch editing zero values using Aspose.Cells
// Tags: batch hide zero cells Aspose.Cells | process multiple .xlsx files C# | iterate worksheets cells Aspose.Cells | clear numeric zero values Excel .NET | skip password protected workbooks Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The C# program scans a specified folder for .xlsx files, loads each workbook with Aspose.Cells, iterates through every worksheet and its used range, clears cells that contain a numeric zero, saves the workbook back to the original file, and gracefully skips password‑protected files.
class Program
{
    static void Main()
    {
        // Path to the folder containing the Excel workbooks
        string folderPath = @"C:\Workbooks";

        // Verify that the folder exists
        if (!Directory.Exists(folderPath))
        {
            Console.WriteLine($"Folder not found: {folderPath}");
            return;
        }

        // Retrieve all .xlsx files in the specified folder
        string[] workbookFiles = Directory.GetFiles(folderPath, "*.xlsx");

        foreach (string filePath in workbookFiles)
        {
            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                continue;
            }

            try
            {
                // Load the workbook (no password). If the file is password‑protected,
                // an exception will be thrown and handled below.
                Workbook workbook = new Workbook(filePath);

                // Iterate through every worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Determine the used range boundaries
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    // Loop through each cell within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];

                            // Identify numeric cells that contain a zero value
                            if (cell.Type == CellValueType.IsNumeric && cell.DoubleValue == 0.0)
                            {
                                // Hide the zero by clearing the cell's content
                                cell.PutValue(string.Empty);
                            }
                        }
                    }
                }

                // Save the modified workbook back to the original file
                workbook.Save(filePath);
                Console.WriteLine($"Processed: {filePath}");
            }
            catch (Exception ex)
            {
                // Detect password‑protected files via exception message
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine($"Skipped password‑protected file: {filePath}");
                }
                else
                {
                    // Log any other errors and continue processing remaining files
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
    }
}
