// Title: Export a Numeric Column to a Little‑Endian Binary File with Aspose.Cells (C#)
// Description: Creates a workbook, fills column A with numbers, finds the last populated row, iterates through each cell in that column, converts numeric values to little‑endian byte arrays using BitConverter, and writes them sequentially to a binary file via BinaryWriter.
// Keywords: Aspose.Cells export column to binary | C# little endian binary file | numeric cells to binary | Excel column serialization | BinaryWriter Aspose.Cells | BitConverter GetBytes double | cell enumeration C# | serialize Excel numeric data | binary file generation from worksheet | compact numeric dataset storage
// Common Searches: how to export numeric column from Aspose.Cells to binary C# | write Excel column values as little endian bytes | Aspose.Cells iterate column and save to binary file | C# convert double to little endian byte array | binary export of worksheet numeric data
// Developer Intent: Extract all numeric values from a specific worksheet column and write them as little‑endian binary data.
// Use Cases: Transmit numeric datasets to low‑level hardware that expects little‑endian doubles | Create a compact binary cache for high‑performance data processing | Serialize column data for custom file formats without using CSV or XML | Share numeric results between .NET applications and native C/C++ modules
// AI Prompts: Generate C# code that uses Aspose.Cells to read a given column, convert each numeric cell to a little‑endian byte array, and append the bytes to a binary file. | Explain why BitConverter.GetBytes returns little‑endian bytes on Windows and how to ensure correct endianness on other platforms. | Modify the example to preserve integer precision by writing ints as 4‑byte little‑endian values and doubles as 8‑byte values in the same binary stream.

using System;
using System.IO;
using Aspose.Cells;

namespace ExportNumericColumnToBinary
{
    // Creates a workbook, fills column A with numbers, finds the last populated row, iterates through each cell in that column, converts numeric values to little‑endian byte arrays using BitConverter, and writes them sequentially to a binary file via BinaryWriter.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample numeric data in column A (index 0)
            for (int i = 0; i < 10; i++)
            {
                // Put some integer and double values
                cells[i, 0].PutValue(i * 10);               // integer values
                cells[i, 1].PutValue(i + 0.5);              // auxiliary column (not exported)
            }

            // Determine the last row that contains data in column A
            int lastRow = cells.MaxDataRow; // zero‑based index

            // Path of the binary file to write
            string binaryFilePath = "numericColumn.bin";

            // Open a binary writer (writes in little‑endian by default)
            using (FileStream fs = new FileStream(binaryFilePath, FileMode.Create, FileAccess.Write))
            using (BinaryWriter writer = new BinaryWriter(fs))
            {
                // Enumerate each cell in column A from row 0 to lastRow
                for (int row = 0; row <= lastRow; row++)
                {
                    Cell cell = cells[row, 0];

                    // Process only numeric cells (int, double, decimal, etc.)
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        // Retrieve the numeric value as double (covers int, float, etc.)
                        double numericValue = cell.DoubleValue;

                        // Convert the double to little‑endian byte array
                        byte[] bytes = BitConverter.GetBytes(numericValue); // BitConverter uses little‑endian on Windows

                        // Write the bytes to the binary file
                        writer.Write(bytes);
                    }
                }
            }

            Console.WriteLine($"Numeric column exported to binary file: {binaryFilePath}");
        }
    }
}
