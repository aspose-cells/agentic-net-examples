using System;
using Aspose.Cells;
using System.IO;

class CsvTrimExample
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Configure load options for CSV
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
        loadOptions.Separator = ',';               // Define column separator
        loadOptions.HasTextQualifier = true;       // Preserve quoted text

        // Load the CSV into a workbook
        Workbook workbook = new Workbook(csvPath, loadOptions);
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Determine the used range of the worksheet
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through each cell and trim whitespace from string values
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsString)
                {
                    string trimmed = cell.StringValue.Trim(); // Remove leading/trailing spaces
                    cell.PutValue(trimmed);
                }
            }
        }

        // Save the cleaned data back to CSV
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.Separator = ',';               // Keep the same separator
        workbook.Save("output_clean.csv", saveOptions);
    }
}