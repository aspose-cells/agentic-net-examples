using Aspose.Cells;
using System;

// Author: Aspose.Cells .NET example – load numeric values as text
class Program
{
    static void Main()
    {
        // Configure load options to prevent conversion of numeric strings to numbers
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            ConvertNumericData = false // keep numeric data as text
        };

        // Load a text‑based file (e.g., CSV or TXT) using the specified options
        Workbook workbook = new Workbook("input.txt", loadOptions);

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the value of cell A1 as a string (will be the original text)
        string cellText = sheet.Cells["A1"].StringValue;
        Console.WriteLine($"A1 (as text): {cellText}");

        // Save the workbook to an Excel file (optional verification)
        workbook.Save("output.xlsx");
    }
}