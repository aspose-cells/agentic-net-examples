using System;
using System.Collections.Generic;
using Aspose.Cells;

class UniqueHeadersDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data: populate the first row with some duplicate headers
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Age");
        cells["C1"].PutValue("Name");   // duplicate
        cells["D1"].PutValue("Email");
        cells["E1"].PutValue("Age");    // duplicate

        // Collect distinct header values from the first row
        HashSet<string> uniqueHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int maxColumn = cells.MaxDataColumn; // last column that contains data
        for (int col = 0; col <= maxColumn; col++)
        {
            string header = cells[0, col].StringValue;
            if (!string.IsNullOrEmpty(header))
            {
                uniqueHeaders.Add(header);
            }
        }

        // Convert the set to a list for further processing if needed
        List<string> headerList = new List<string>(uniqueHeaders);

        // Output the unique headers
        Console.WriteLine("Unique column headers:");
        foreach (string header in headerList)
        {
            Console.WriteLine(header);
        }

        // Save the workbook (optional)
        workbook.Save("UniqueHeadersDemo.xlsx");
    }
}