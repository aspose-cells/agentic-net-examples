using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class StandardHeightAndFormulaAutoFitDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Input workbook path
            string inputPath = @"C:\Temp\InputWorkbook.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set default row height (points)
            worksheet.Cells.StandardHeight = 20;

            // Determine the range of rows and columns that contain data
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            // Auto‑fit rows that contain at least one formula
            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                bool hasFormula = false;

                for (int colIndex = 0; colIndex <= maxCol; colIndex++)
                {
                    Cell cell = worksheet.Cells[rowIndex, colIndex];
                    if (cell != null && cell.IsFormula)
                    {
                        hasFormula = true;
                        break;
                    }
                }

                if (hasFormula)
                {
                    worksheet.AutoFitRow(rowIndex);
                }
            }

            // Output workbook path
            string outputPath = @"C:\Temp\OutputWorkbook.xlsx";

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}