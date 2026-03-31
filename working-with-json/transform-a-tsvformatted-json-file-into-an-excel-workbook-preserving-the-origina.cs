using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the TSV‑formatted JSON file
        string tsvFilePath = "input.tsv";

        // Path for the generated Excel workbook
        string outputExcelPath = "output.xlsx";

        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure load options for a tab‑separated file
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            Separator = '\t',          // Tab delimiter
            ConvertNumericData = true, // Preserve numeric types
            ConvertDateTimeData = true // Preserve date/time types
        };

        // Import the TSV data starting at cell A1 (row 0, column 0)
        // (lifecycle: load)
        cells.ImportCSV(tsvFilePath, loadOptions, 0, 0);

        // Save the workbook as XLSX (lifecycle: save)
        workbook.Save(outputExcelPath, SaveFormat.Xlsx);
    }
}