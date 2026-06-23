using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Sample CSV content
        string csvData = "Header1,Header2,Header3\n" +
                         "Data1,Data2,Data3\n" +
                         "Data4,Data5,Data6\n" +
                         "Data7,Data8,Data9";

        // Define load options with the desired number of header rows
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            HeaderRowsCount = 1 // Number of rows to treat as header
        };

        // Store the header rows count for later use
        int headerRowsToFreeze = loadOptions.HeaderRowsCount;

        // Load the CSV into a workbook using the specified options
        using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
        {
            Workbook workbook = new Workbook(stream, loadOptions);
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze the header rows in the worksheet
            // row and freezedRows parameters are set to the header count
            worksheet.FreezePanes(headerRowsToFreeze, 0, headerRowsToFreeze, 0);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("Result.xlsx");
        }

        // Output the stored header rows count
        Console.WriteLine("Header rows frozen: " + headerRowsToFreeze);
    }
}