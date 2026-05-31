using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBinaryExport
{
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

                // Populate numeric data in column A (index 0)
                for (int i = 0; i < 10; i++)
                {
                    // Example values: 1.0, 2.0, ..., 10.0
                    cells[i, 0].PutValue(i + 1);
                }

                // Determine the last row that contains data in the column (zero‑based)
                int lastRow = cells.MaxDataRow;

                // Path of the binary file to write
                string binaryFilePath = "NumericColumnData.bin";

                // Write numeric values to a binary file (little‑endian)
                using (FileStream fs = new FileStream(binaryFilePath, FileMode.Create, FileAccess.Write))
                using (BinaryWriter writer = new BinaryWriter(fs))
                {
                    for (int row = 0; row <= lastRow; row++)
                    {
                        Cell cell = cells[row, 0];

                        // Process only numeric cells
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            double value = cell.DoubleValue;
                            byte[] bytes = BitConverter.GetBytes(value);

                            // Ensure little‑endian order regardless of platform
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
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}