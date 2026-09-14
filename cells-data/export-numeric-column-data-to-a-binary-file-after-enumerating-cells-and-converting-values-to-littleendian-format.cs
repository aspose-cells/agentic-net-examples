// Title: Export a numeric column from an Aspose.Cells worksheet to a little‑endian binary file in C#
// AI Prompts: Write C# code that iterates through a specific column in an Aspose.Cells worksheet, extracts only numeric cells, converts each double to a little‑endian byte array, and writes the bytes sequentially to a .bin file. | Create a reusable method ExportColumnToBinary(Worksheet sheet, int columnIndex, string outputPath) that saves all numeric values from the given column as little‑endian doubles using BinaryWriter. | Extend the export routine to prepend a 4‑byte little‑endian integer indicating the total count of numeric values before writing the column data.
// Common Searches: Aspose.Cells C# export column values to binary file with little endian format | How to write Excel numeric column to .bin using Aspose.Cells and .NET | C# iterate over worksheet column and save doubles as little endian bytes | BinaryWriter example for exporting Excel numeric data with Aspose.Cells
// Tags: Aspose.Cells numeric column binary export | C# little-endian double serialization | BinaryWriter Excel data extraction | export worksheet column to .bin file | iterate cells numeric type Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExportNumericColumnToBinary
{
    // The sample creates a workbook, fills column A with numbers 1‑10, then walks that column, converts each numeric cell to a little‑endian double byte array, and writes the bytes sequentially to NumericColumnData.bin using BinaryWriter.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate column A (index 0) with numeric values 1..10
                for (int row = 0; row < 10; row++)
                {
                    cells[row, 0].PutValue(row + 1); // A1..A10
                }

                // Path of the binary file to write
                string binaryFilePath = "NumericColumnData.bin";

                // Ensure the directory for the binary file exists
                string directory = Path.GetDirectoryName(binaryFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Open a file stream for writing binary data
                using (FileStream fs = new FileStream(binaryFilePath, FileMode.Create, FileAccess.Write))
                using (BinaryWriter writer = new BinaryWriter(fs))
                {
                    // Determine the last row that contains data in column 0
                    int lastRow = cells.MaxDataRow;

                    // Enumerate each cell in the numeric column
                    for (int row = 0; row <= lastRow; row++)
                    {
                        Cell cell = cells[row, 0];

                        // Process only numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            double numericValue = cell.DoubleValue;

                            // Convert the double to a byte array (little‑endian)
                            byte[] bytes = BitConverter.GetBytes(numericValue);
                            if (!BitConverter.IsLittleEndian)
                            {
                                Array.Reverse(bytes);
                            }

                            // Write the bytes to the binary file
                            writer.Write(bytes);
                        }
                    }
                }

                Console.WriteLine($"Numeric column data has been exported to '{binaryFilePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
