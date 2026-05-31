using System;
using System.IO;
using Aspose.Cells;

class DeleteRowsWithNullInRequiredColumn
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Index of the required column (0 = column A)
            int requiredColumnIndex = 0;

            // Iterate from the last data row up to the first row
            // Deleting from bottom prevents index shifting issues
            for (int row = cells.MaxDataRow; row >= 0; row--)
            {
                Cell cell = cells[row, requiredColumnIndex];

                // Determine if the cell is considered "null" (blank, DBNull, or empty string)
                bool isNull = cell.Value == null ||
                              (cell.Type == CellValueType.IsString && string.IsNullOrWhiteSpace(cell.StringValue));

                if (isNull)
                {
                    // Delete the entire row
                    cells.DeleteRow(row);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}