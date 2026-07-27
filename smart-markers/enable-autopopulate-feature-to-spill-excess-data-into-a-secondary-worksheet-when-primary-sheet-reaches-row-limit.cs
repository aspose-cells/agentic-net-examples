using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class AutoPopulateExample
{
    static void Main()
    {
        // Generate CSV data that exceeds the maximum rows per worksheet (1,048,576)
        const int totalRows = 1_050_000; // 1,050,000 rows will force overflow
        var sb = new StringBuilder();

        // Header row
        sb.AppendLine("Index,Value");

        // Populate rows
        for (int i = 1; i <= totalRows; i++)
        {
            sb.AppendLine($"{i},Data_{i}");
        }

        // Convert CSV string to a memory stream (UTF‑8 encoding)
        using var csvStream = new MemoryStream(Encoding.UTF8.GetBytes(sb.ToString()));

        // Enable the auto‑populate (spill‑over) feature
        var loadOptions = new TxtLoadOptions
        {
            ExtendToNextSheet = true   // When row limit is hit, data continues on a new sheet
        };

        // Load the CSV data into a workbook using the specified options
        var workbook = new Workbook(csvStream, loadOptions);

        // Optional: display information about the created workbook
        Console.WriteLine($"Worksheets created: {workbook.Worksheets.Count}");
        Console.WriteLine($"Rows in first sheet: {workbook.Worksheets[0].Cells.MaxDataRow + 1}");
        if (workbook.Worksheets.Count > 1)
        {
            Console.WriteLine($"Rows in second sheet: {workbook.Worksheets[1].Cells.MaxDataRow + 1}");
        }

        // Save the workbook to an XLSX file
        workbook.Save("SpillOverResult.xlsx", SaveFormat.Xlsx);
    }
}