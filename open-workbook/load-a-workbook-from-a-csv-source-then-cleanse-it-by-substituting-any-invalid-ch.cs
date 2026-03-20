using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

class CsvCleanseDemo
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Get the Cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Import CSV data (comma delimiter, convert numeric data)
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Define characters that are not allowed in Excel cell values
        char[] invalidChars = new char[] { '*', '?', ':', '[', ']', '/', '\\' };

        // Iterate through all populated cells and cleanse their string values
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsString)
                {
                    string value = cell.StringValue;

                    // Replace each invalid character with a space
                    foreach (char ch in invalidChars)
                    {
                        value = value.Replace(ch, ' ');
                    }

                    // Remove control characters (ASCII < 32)
                    value = Regex.Replace(value, @"[\x00-\x1F]", " ");

                    // Write the cleansed value back to the cell
                    cell.PutValue(value);
                }
            }
        }

        // Save the cleansed workbook to an XLSX file
        workbook.Save("cleansed_output.xlsx", SaveFormat.Xlsx);
    }
}