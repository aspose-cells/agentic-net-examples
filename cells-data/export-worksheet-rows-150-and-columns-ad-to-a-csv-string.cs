using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class ExportCsvExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // OPTIONAL: Populate some sample data for demonstration
        for (int row = 0; row < 60; row++)          // more than 50 rows
        {
            for (int col = 0; col < 6; col++)       // more than 4 columns
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define the export area: rows 1‑50 (0‑49) and columns A‑D (0‑3)
        CellArea exportArea = new CellArea
        {
            StartRow = 0,
            EndRow = 49,
            StartColumn = 0,
            EndColumn = 3
        };

        // Configure TxtSaveOptions for CSV output
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.Separator = ',';          // CSV delimiter
        saveOptions.ExportArea = exportArea;   // limit export to the defined range

        // Save the selected range to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, saveOptions);

            // Convert the memory stream to a CSV string
            string csvString = Encoding.UTF8.GetString(ms.ToArray());

            // Output the CSV string
            Console.WriteLine(csvString);
        }
    }
}