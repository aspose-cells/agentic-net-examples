// Title: Export numeric values from a specific Excel column to a little‑endian binary file using Aspose.Cells for .NET
// Common Searches: how to write numeric column data from an .xlsx file to a binary file with Aspose.Cells in C# | convert Excel double values to little‑endian byte array using Aspose.Cells | enumerate cells in a column and save them as binary using Aspose.Cells .NET | export first worksheet column numeric values to a binary stream C# Aspose.Cells | handle password‑protected workbook while exporting column data to binary
// Tags: Aspose.Cells column enumeration to binary | Aspose.Cells write double little-endian | Aspose.Cells load workbook password handling | Aspose.Cells binary writer for numeric data | Aspose.Cells convert double to little-endian bytes

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook (with optional password handling), selects the first worksheet, iterates through the used rows of a chosen column, extracts numeric cells, converts each double to a little‑endian byte array, and writes the bytes sequentially to a binary file using a BinaryWriter.
class Program
{
    static void Main(string[] args)
    {
        // Paths – adjust as needed
        string excelPath = "input.xlsx";
        string binaryPath = "output.bin";

        try
        {
            // Verify that the input Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.Error.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook (use LoadOptions in case the file is password‑protected)
            Workbook workbook;
            try
            {
                var loadOptions = new LoadOptions(); // no password supplied
                workbook = new Workbook(excelPath, loadOptions);
            }
            catch (CellsException ex)
            {
                // Aspose.Cells throws CellsException for password‑protected files; check the message
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.Error.WriteLine("Error: The workbook is password‑protected. Provide a valid password.");
                    return;
                }
                // Re‑throw other CellsException types
                throw;
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the column to export (0‑based index, e.g., column A = 0)
            int columnIndex = 0;

            // Ensure the column index is within the worksheet's column range
            if (columnIndex < 0 || columnIndex > sheet.Cells.MaxColumn)
            {
                Console.Error.WriteLine($"Error: Column index {columnIndex} is out of range.");
                return;
            }

            // Prepare a binary writer for the output file
            using (FileStream fs = new FileStream(binaryPath, FileMode.Create, FileAccess.Write))
            using (BinaryWriter writer = new BinaryWriter(fs))
            {
                // Determine the used range for the column
                int firstRow = sheet.Cells.MinRow;
                int lastRow = sheet.Cells.MaxRow;

                // Enumerate each cell in the column within the used range
                for (int row = firstRow; row <= lastRow; row++)
                {
                    Cell cell = sheet.Cells[row, columnIndex];

                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double value = cell.DoubleValue;

                        // Convert the double to little‑endian byte array
                        byte[] bytes = BitConverter.GetBytes(value);
                        if (!BitConverter.IsLittleEndian)
                        {
                            // Reverse to little‑endian if running on a big‑endian platform
                            Array.Reverse(bytes);
                        }

                        // Write the bytes to the binary file
                        writer.Write(bytes);
                    }
                }
            }

            // Inform the user
            Console.WriteLine($"Numeric data from column {columnIndex} exported to {binaryPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
