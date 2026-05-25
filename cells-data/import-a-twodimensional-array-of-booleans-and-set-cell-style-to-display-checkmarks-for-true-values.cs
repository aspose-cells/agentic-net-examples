using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCheckBoxDemo
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

                // Define a two‑dimensional boolean array
                bool[,] boolData = new bool[,]
                {
                    { true,  false, true  },
                    { false, true,  false },
                    { true,  true,  false }
                };

                int rows = boolData.GetLength(0);
                int cols = boolData.GetLength(1);

                // Convert bool[,] to object[,] required by ImportTwoDimensionArray
                object[,] objData = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        objData[r, c] = boolData[r, c];
                    }
                }

                // Import the boolean data starting at cell A1 (row 0, column 0)
                cells.ImportTwoDimensionArray(objData, 0, 0);

                // After import, set each cell to checkbox style while preserving its value
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        Cell cell = cells[r, c];
                        cell.IsCheckBoxStyle = true; // Enable checkbox visual
                    }
                }

                // Define output file path
                string outputPath = "BooleanCheckBoxDemo.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}