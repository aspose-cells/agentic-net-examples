using System;
using Aspose.Cells;

class SplitAddressExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Import the CSV file (assumed to be comma‑separated)
        string csvPath = "input.csv";
        // Parameters: file name, delimiter, convert numeric data, first row, first column
        cells.ImportCSV(csvPath, ",", true, 0, 0);

        // Index of the combined address column (zero‑based). Adjust as needed.
        int addressColumnIndex = 2; // e.g., column C

        // Configure TextToColumns options to split by comma
        TxtLoadOptions txtOptions = new TxtLoadOptions();
        txtOptions.Separator = ',';

        // Number of rows to process (including header if present)
        int totalRows = cells.MaxDataRow + 1;

        // Split the address column into separate columns (street, city, zip)
        cells.TextToColumns(0, addressColumnIndex, totalRows, txtOptions);

        // Save the workbook with the split columns
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}