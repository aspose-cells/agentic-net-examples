using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class LoadCsvExample
{
    static void Main()
    {
        // Path to the CSV file that may contain special characters
        string csvPath = "data.csv";

        // Create a new empty workbook
        Workbook workbook = new Workbook();

        // Get the Cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Configure CSV load options
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Encoding = Encoding.UTF8;          // Correct encoding for special characters
        loadOptions.Separator = ',';                  // Standard comma separator
        loadOptions.HasTextQualifier = true;          // Enable text qualifier handling
        loadOptions.TextQualifier = '"';              // Use double quote as text qualifier

        // Open the CSV file as a stream and import its content starting at cell A1 (row 0, column 0)
        using (FileStream stream = new FileStream(csvPath, FileMode.Open, FileAccess.Read))
        {
            cells.ImportCSV(stream, loadOptions, 0, 0);
        }

        // Save the workbook to an Excel file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}