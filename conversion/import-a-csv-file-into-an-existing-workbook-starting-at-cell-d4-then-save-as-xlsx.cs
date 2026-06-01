using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the existing workbook that will receive the CSV data
        string workbookPath = "ExistingWorkbook.xlsx";

        // Path to the CSV file to be imported
        string csvPath = "DataFile.csv";

        // Path for the resulting XLSX file
        string outputPath = "ResultWorkbook.xlsx";

        // Load the existing workbook (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(workbookPath);

        // Get the first worksheet and its Cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Import the CSV file starting at cell D4 (row index 3, column index 3)
        // Using comma as the delimiter and converting numeric strings to numbers
        cells.ImportCSV(csvPath, ",", true, 3, 3);

        // Save the modified workbook as XLSX (uses Workbook.Save method)
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}