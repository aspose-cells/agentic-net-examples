// Title: How to auto‑fit columns C to F in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that populates cells C1:F2, calls Worksheet.AutoFitColumns(2, 4) to auto‑size those columns, and saves the workbook as an .xlsx file. | Explain how to calculate the start column index and column count arguments for Worksheet.AutoFitColumns when targeting a specific column range. | Write a reusable C# method that accepts a worksheet, a start column index, and a column count, applies AutoFitColumns, and returns the adjusted worksheet.
// Common Searches: Aspose.Cells C# auto fit columns C-F example code | Worksheet.AutoFitColumns start column and count parameters usage | How to auto size a specific range of columns in an Excel file with Aspose.Cells .NET | C# adjust column width for columns 3 to 6 using Aspose.Cells
// Tags: auto-fit specific column range Aspose.Cells | Worksheet.AutoFitColumns start index parameter | C# column width adjustment Aspose.Cells | Excel column auto sizing .NET | auto-fit columns range C-F Aspose.Cells

using Aspose.Cells;
using System;

// // This C# program creates a new workbook, writes sample data to columns C‑F, auto‑fits those columns with Worksheet.AutoFitColumns(2, 4), and saves the file as AutoFitColumnsExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data in columns C-F
        sheet.Cells["C1"].PutValue("Header C");
        sheet.Cells["D1"].PutValue("Header D");
        sheet.Cells["E1"].PutValue("Header E");
        sheet.Cells["F1"].PutValue("Header F");
        sheet.Cells["C2"].PutValue("Longer text in column C");
        sheet.Cells["D2"].PutValue("Data D");
        sheet.Cells["E2"].PutValue("Data E");
        sheet.Cells["F2"].PutValue("Data F");

        // Auto‑fit columns C (index 2) through F (index 5)
        // startColumn = 2, totalColumns = 4 (C, D, E, F)
        sheet.AutoFitColumns(2, 4);

        // Save the workbook
        workbook.Save("AutoFitColumnsExample.xlsx");
    }
}
