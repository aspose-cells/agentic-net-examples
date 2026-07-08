using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in rows 1‑50 and columns A‑D (indices 0‑49, 0‑3)
        for (int row = 0; row < 50; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Configure TxtSaveOptions to export only the desired area (A1:D50)
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);
        saveOptions.ExportArea = new CellArea
        {
            StartRow = 0,      // Row 1 (zero‑based)
            EndRow = 49,       // Row 50
            StartColumn = 0,   // Column A
            EndColumn = 3      // Column D
        };
        // Optional: set the CSV separator (default is comma)
        saveOptions.Separator = ',';

        // Save the selected range to a memory stream in CSV format
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, saveOptions);
            // Convert the stream content to a CSV string
            string csvString = Encoding.UTF8.GetString(stream.ToArray());

            // Output the CSV string
            Console.WriteLine(csvString);
        }
    }
}