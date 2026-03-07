using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImportTrimSpaces
{
    class Program
    {
        static void Main()
        {
            // Create a sample CSV file containing extra spaces around values
            string csvPath = "sample_with_spaces.csv";
            string csvContent = "Name , Age , City\n John , 30 , New York\n Alice , 25 , London";
            File.WriteAllText(csvPath, csvContent);

            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Prepare TxtLoadOptions for CSV import
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                // Use comma as separator
                Separator = ',',
                // Treat consecutive delimiters as one (helps when spaces are present)
                TreatConsecutiveDelimitersAsOne = true,
                // Convert numeric strings to numbers
                ConvertNumericData = true
            };

            // Import the CSV file starting at cell A1 (row 0, column 0)
            // Lifecycle rule: load (ImportCSV with TxtLoadOptions)
            workbook.Worksheets[0].Cells.ImportCSV(csvPath, loadOptions, 0, 0);

            // Trim extra spaces from all string cells after import
            Cells cells = workbook.Worksheets[0].Cells;
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
                        // Update the cell only if trimming changes the value
                        if (trimmed != cell.StringValue)
                        {
                            cell.PutValue(trimmed);
                        }
                    }
                }
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("TrimmedOutput.xlsx", SaveFormat.Xlsx);

            // Clean up the temporary CSV file
            File.Delete(csvPath);
        }
    }
}