// Title: How to split comma‑separated strings in column B into separate columns using Aspose.Cells TextToColumns in C#
// AI Prompts: Generate C# code that configures TxtLoadOptions with a comma separator and applies Cells.TextToColumns to split the values in column B of a worksheet into individual columns. | Show a complete example that creates a workbook, writes CSV strings to column B, runs TextToColumns with a custom delimiter, captures the resulting column count, and saves the file as an .xlsx. | Explain how to retrieve and use the number of columns produced after the TextToColumns operation in Aspose.Cells.
// Common Searches: asp.net core split column B csv values into separate cells using Aspose.Cells TextToColumns | c# Aspose.Cells TextToColumns delimiter comma example | how to convert comma separated strings in an Excel column to multiple columns with Aspose.Cells | using TxtLoadOptions separator property to split Excel column data in C#
// Tags: Aspose.Cells TextToColumns comma delimiter | C# TxtLoadOptions separator property | split Excel column CSV strings Aspose.Cells | convert column B to multiple columns C# | save workbook as XLSX after TextToColumns

using System;
using Aspose.Cells;

// The sample creates a new workbook, writes comma‑separated strings into cells B1‑B3, sets TxtLoadOptions.Separator to a comma, calls Cells.TextToColumns starting at row 0 column 1 for three rows, captures the total number of columns generated, and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column B (index 1) with comma‑separated values
        cells["B1"].PutValue("John,Doe,30");
        cells["B2"].PutValue("Jane,Smith,28");
        cells["B3"].PutValue("Bob,Johnson,45");

        // Configure TextLoadOptions to use a comma as the delimiter
        TxtLoadOptions options = new TxtLoadOptions();
        options.Separator = ',';

        // Split the text in column B into separate columns
        // Parameters: start row (0), start column (1), number of rows (3), options
        int totalColumns = cells.TextToColumns(0, 1, 3, options);

        Console.WriteLine($"Total columns after split: {totalColumns}");

        // Save the workbook to an Excel file
        workbook.Save("output.xlsx");
    }
}
