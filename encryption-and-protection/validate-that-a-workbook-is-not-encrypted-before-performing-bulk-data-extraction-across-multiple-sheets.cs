// Title: Detect Excel encryption and read every sheet with Aspose.Cells for .NET
// Description: A C# sample that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to identify password‑protected workbooks, skips encrypted files, then loads the workbook and iterates through all worksheets, outputting each populated cell's address and value.
// Keywords: Aspose.Cells encryption detection | FileFormatUtil DetectFileFormat C# | read non‑encrypted Excel .NET | bulk worksheet data extraction | skip password‑protected workbook | iterate all cells Aspose.Cells | Excel file validation before import
// Common Searches: How to check if an Excel file is password protected using Aspose.Cells | C# detect encrypted workbook without opening it | Extract data from all sheets after confirming workbook is not encrypted | Skip encrypted Excel files in a batch process Aspose.Cells | FileFormatUtil DetectFileFormat example
// Developer Intent: Confirm that an Excel file is not password‑protected before extracting data from every worksheet.
// Use Cases: Validate user‑uploaded Excel files on a server and reject encrypted ones before import. | Perform a full data dump of all worksheets in a non‑encrypted workbook for migration or reporting. | Run a scheduled batch job that processes many Excel files, automatically bypassing any that are encrypted.
// AI Prompts: Generate a reusable C# method that uses Aspose.Cells to detect workbook encryption and, if clear, returns a dictionary of sheet names mapped to lists of non‑empty cell addresses and values. | Create robust error‑handling code for a bulk Excel processing pipeline that logs encrypted files and continues with the remaining files using FileFormatUtil. | Refactor the provided example into two separate functions—one for encryption validation and one for data extraction—while preserving async support.

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookEncryptionValidatorApp
{
    // A C# sample that uses Aspose.Cells' FileFormatUtil.DetectFileFormat to identify password‑protected workbooks, skips encrypted files, then loads the workbook and iterates through all worksheets, outputting each populated cell's address and value.
    public class WorkbookEncryptionValidator
    {
        // Validates that the workbook is not encrypted and extracts data from all sheets.
        public static void ExtractData(string filePath)
        {
            try
            {
                // Verify that the file exists before attempting to load it.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Detect file format and encryption status without loading the workbook.
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
                if (formatInfo.IsEncrypted)
                {
                    Console.WriteLine($"The workbook \"{filePath}\" is encrypted. Extraction aborted.");
                    return;
                }

                // Load the workbook (no password required because it is not encrypted).
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Iterate through each worksheet in the workbook.
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        Console.WriteLine($"--- Sheet: {sheet.Name} ---");

                        // Determine the used range.
                        int maxRow = sheet.Cells.MaxDataRow;
                        int maxCol = sheet.Cells.MaxDataColumn;

                        // Loop through all used cells and output their values.
                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                Cell cell = sheet.Cells[row, col];
                                if (cell.Value != null)
                                {
                                    // Row and column numbers are 1‑based for readability.
                                    Console.WriteLine($"R{row + 1}C{col + 1}: {cell.Value}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    class Program
    {
        // Entry point of the application.
        static void Main(string[] args)
        {
            string filePath;

            if (args.Length > 0)
            {
                filePath = args[0];
            }
            else
            {
                Console.Write("Enter the path to the Excel file: ");
                filePath = Console.ReadLine();
            }

            WorkbookEncryptionValidator.ExtractData(filePath);
        }
    }
}
