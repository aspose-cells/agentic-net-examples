// Title: How to programmatically verify SUMPRODUCT support in Aspose.Cells before using the formula in C#
// Common Searches: how to check if SUMPRODUCT is supported by Aspose.Cells in .NET | list of Excel functions available in Aspose.Cells C# API | validate Excel formula compatibility with Aspose.Cells before saving workbook | prevent runtime error caused by unsupported Excel functions in Aspose.Cells | C# code to query supported functions in Aspose.Cells library
// Tags: Aspose.Cells supported functions list | SUMPRODUCT support verification Aspose.Cells | C# Aspose.Cells formula compatibility check | runtime validation of Excel formulas Aspose.Cells | handle unsupported Excel function exception Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, fills cells A1‑A3 and B1‑B3 with numeric data, assigns a SUMPRODUCT formula to C1, and saves the file as Output.xlsx. It demonstrates wrapping the operation in a try‑catch block and advises checking the Aspose.Cells supported functions list beforehand to avoid unsupported‑function errors.
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
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["A3"].PutValue(3);
            sheet.Cells["B1"].PutValue(4);
            sheet.Cells["B2"].PutValue(5);
            sheet.Cells["B3"].PutValue(6);

            // Apply SUMPRODUCT formula (supported by Aspose.Cells)
            sheet.Cells["C1"].Formula = "=SUMPRODUCT(A1:A3,B1:B3)";

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
