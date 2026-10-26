// Title: Export numeric column from Excel to a little‑endian binary file using Aspose.Cells for .NET
// Common Searches: how to write Excel numeric column values to a binary file with Aspose.Cells C# | export double values from a specific Excel column to little endian binary using Aspose.Cells | iterate through cells in a column and save numeric data as binary in .NET | Aspose.Cells export column data to binary file little endian | C# read numeric cells from Excel and write to binary using BinaryWriter
// Tags: numeric column binary export Aspose.Cells | little endian double serialization C# | filter numeric cells while iterating Excel column | binarywriter usage with Aspose.Cells extracted data | byte array conversion for double values little endian

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, selects a worksheet, iterates over all rows in a specified column, extracts numeric cells, converts each double to a little‑endian byte array, and writes the bytes sequentially to a binary file using BinaryWriter, handling missing files and password‑protected workbooks.
class ExportNumericColumnToBinary
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Path to the destination binary file
            string binaryPath = "output.bin";

            // Zero‑based index of the column to export (e.g., 0 = column A)
            int columnIndex = 0;

            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Input file not found: {excelPath}");
                return;
            }

            // Load the workbook (no password supplied; handle password‑protected files in catch)
            Workbook workbook = new Workbook(excelPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Open a binary writer for the output file
            using (BinaryWriter writer = new BinaryWriter(File.Open(binaryPath, FileMode.Create)))
            {
                // Iterate through rows that contain data
                for (int row = cells.MinRow; row <= cells.MaxRow; row++)
                {
                    Cell cell = cells[row, columnIndex];

                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double value = cell.DoubleValue;
                        byte[] bytes = BitConverter.GetBytes(value);

                        // Ensure little‑endian order
                        if (!BitConverter.IsLittleEndian)
                        {
                            Array.Reverse(bytes);
                        }

                        writer.Write(bytes);
                    }
                }
            }

            Console.WriteLine("Numeric column data exported to binary file successfully.");
        }
        catch (CellsException cex) when (cex.Message.Contains("password"))
        {
            Console.WriteLine("The workbook is password protected. Provide a password via LoadOptions if required.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
