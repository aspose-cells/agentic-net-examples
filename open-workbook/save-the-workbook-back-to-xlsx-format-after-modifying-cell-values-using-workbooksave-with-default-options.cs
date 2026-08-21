// Title: Save a Modified Workbook to XLSX with Aspose.Cells (C#) Using Default Options
// Description: Demonstrates how to create a new Workbook, write text, a date, and a number to cells A1, B1, A2, and B2, and then persist the changes by calling Workbook.Save with SaveFormat.Xlsx and the library's default settings.
// Keywords: Aspose.Cells C# save workbook | Workbook.Save default options | export to XLSX Aspose.Cells | modify cells Aspose.Cells | C# Excel file generation
// Common Searches: Aspose.Cells save workbook as XLSX C# | How to use Workbook.Save with default settings | C# write values to Excel and save with Aspose | Save modified spreadsheet using Aspose.Cells
// Developer Intent: Persist a workbook after updating cell values without specifying custom save parameters.
// Use Cases: Generate a quick report by filling header and data cells and exporting to XLSX. | Update an existing spreadsheet with the current timestamp and an identifier, then save the result. | Create a simple data dump from a .NET application and store it as a standard Excel file.
// AI Prompts: Show C# code that opens an existing XLSX file, changes several cells, and saves it without overwriting the original. | Explain how to customize Workbook.Save with SaveOptions for compression, password protection, or macro preservation. | Provide examples of converting a workbook to PDF, CSV, and HTML using Aspose.Cells in C#.

using System;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, write text, a date, and a number to cells A1, B1, A2, and B2, and then persist the changes by calling Workbook.Save with SaveFormat.Xlsx and the library's default settings.
class Program
{
    static void Main()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Modify cell values
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");
        worksheet.Cells["A2"].PutValue(DateTime.Now);
        worksheet.Cells["B2"].PutValue(12345);

        // Save the workbook back to XLSX format using default save options
        workbook.Save("ModifiedWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
