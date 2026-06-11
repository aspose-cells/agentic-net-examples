using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Sample CSV data
        string csvData = "Header1,Header2,Header3,Data1,Data2\n" +
                         "H1,H2,H3,D1,D2\n" +
                         "A,B,C,D,E\n" +
                         "G,H,I,J,K";

        // Configure TxtLoadOptions with the desired number of header columns
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            HeaderColumnsCount = 3,          // Number of columns to treat as headers
            Encoding = Encoding.UTF8
        };

        // Load the CSV into a workbook using the options
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(stream, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the number of header columns to freeze
            int headerColumnsToFreeze = loadOptions.HeaderColumnsCount;

            // Freeze only the header columns (no rows are frozen)
            // FreezePanes(rowIndex, columnIndex, freezedRows, freezedColumns)
            worksheet.FreezePanes(0, headerColumnsToFreeze, 0, headerColumnsToFreeze);

            // Store the count for later use (example: write to a cell)
            int storedHeaderCount = headerColumnsToFreeze;
            worksheet.Cells["A1"].PutValue($"Header columns frozen: {storedHeaderCount}");

            // Save the workbook
            workbook.Save("HeaderColumnsFreezeDemo.xlsx");
        }
    }
}