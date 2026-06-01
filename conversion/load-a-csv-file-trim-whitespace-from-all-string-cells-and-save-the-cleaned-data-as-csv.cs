using System;
using Aspose.Cells;

namespace CsvTrimExample
{
    class Program
    {
        static void Main()
        {
            // Paths for input and output CSV files
            string inputCsvPath = "input.csv";
            string outputCsvPath = "output.csv";

            // Create an empty workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells collection
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Import the CSV data starting at cell A1 (row 0, column 0)
            // Use comma as delimiter, do not convert numeric data so all values stay as strings initially
            cells.ImportCSV(inputCsvPath, ",", false, 0, 0);

            // Trim whitespace from all string cells
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string trimmed = cell.StringValue.Trim();
                        cell.PutValue(trimmed);
                    }
                }
            }

            // Save the cleaned data back to CSV
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ','
            };
            workbook.Save(outputCsvPath, saveOptions);
        }
    }
}