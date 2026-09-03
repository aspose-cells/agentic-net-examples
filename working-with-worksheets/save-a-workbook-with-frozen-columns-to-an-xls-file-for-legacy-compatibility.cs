// Title: Create an Excel 97‑2003 (.xls) workbook with frozen columns using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, freezes the first two columns, and saves it as an .xls file with Aspose.Cells. | Adjust the example to freeze a configurable number of columns before exporting to the Excel97To2003 format. | Add comprehensive error handling and output the full file system path of the saved .xls workbook.
// Common Searches: Aspose.Cells C# freeze first two columns and save as Excel 97-2003 file | how to export a workbook with frozen panes to .xls using Aspose.Cells | C# example for FreezePanes and SaveFormat.Excel97To2003 in Aspose.Cells | legacy .xls generation with column freezing in Aspose.Cells for .NET
// Tags: freeze panes Aspose.Cells C# | save workbook as Excel97To2003 Aspose.Cells | frozen columns legacy .xls generation | Aspose.Cells FreezePanes example | C# export workbook to .xls with frozen columns

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a new workbook, adds sample data, freezes the first two columns (A and B), and saves the file as an Excel 97‑2003 (.xls) workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["C1"].PutValue("Score");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["C2"].PutValue(85);
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["C3"].PutValue(92);

                // Freeze the first two columns (A and B)
                // Parameters: totalRows, totalColumns, row, column
                sheet.FreezePanes(0, 2, 0, 2);

                // Save the workbook as an XLS file for legacy compatibility
                string outputFile = "FrozenColumnsWorkbook.xls";
                workbook.Save(outputFile, SaveFormat.Excel97To2003);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
