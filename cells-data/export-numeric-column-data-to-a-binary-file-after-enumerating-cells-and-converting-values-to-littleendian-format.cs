using System;
using System.IO;
using Aspose.Cells;

namespace ExportNumericColumnToBinary
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with numeric values (including a header)
            cells["A1"].PutValue("Header"); // non‑numeric, will be skipped
            for (int i = 0; i < 10; i++)
            {
                // Rows are zero‑based; row 1 corresponds to cell A2
                cells[i + 1, 0].PutValue(i * 1.5); // example numeric data
            }

            // Path of the binary file to write
            string binaryFilePath = "numericColumn.bin";

            // Open a binary writer (writes in little‑endian by default)
            using (FileStream fs = new FileStream(binaryFilePath, FileMode.Create, FileAccess.Write))
            using (BinaryWriter writer = new BinaryWriter(fs))
            {
                // Iterate through all used rows in column A
                int maxRow = cells.MaxDataRow; // last row that contains data
                for (int row = 0; row <= maxRow; row++)
                {
                    Cell cell = cells[row, 0]; // column 0 = A
                    // Process only numeric cells
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        double numericValue = cell.DoubleValue;
                        // Convert to little‑endian byte array (BitConverter uses little‑endian on Windows)
                        byte[] bytes = BitConverter.GetBytes(numericValue);
                        // Ensure little‑endian order (swap if running on big‑endian platform)
                        if (!BitConverter.IsLittleEndian)
                        {
                            Array.Reverse(bytes);
                        }
                        writer.Write(bytes);
                    }
                }
            }

            Console.WriteLine($"Numeric column data exported to binary file: {binaryFilePath}");
        }
    }
}