// Title: Copy cell style with built‑in currency number format using Aspose.Cells CopyStyle in C#
// AI Prompts: Use Aspose.Cells CopyStyle to transfer the currency number format from cell A1 to cell B2 in a .NET workbook. | Demonstrate preserving built‑in number formatting when copying styles between ranges with C# and Aspose.Cells. | Show how to apply a source cell’s style, including its currency format, to a destination range using the CopyStyle method.
// Common Searches: aspnet copy style retain currency format aspocells | c# aspocells copystyle keep number format | how to preserve built‑in number format when copying cell style in Aspose.Cells | copy cell style with currency formatting using Aspose.Cells .NET
// Tags: CopyStyle method currency number format | preserve number formatting Aspose.Cells | copy cell style between ranges C# | built‑in number format inheritance Aspose.Cells | Excel workbook style copy Aspose.Cells .NET

using System;
using Aspose.Cells;

// Alias to avoid conflict with System.Range
using CellsRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // The example creates a workbook, applies a built‑in currency number format (ID 5) to cell A1, copies the entire style—including the number format—to range B2 using the CopyStyle method, inserts a value into B2 to demonstrate the inherited formatting, and saves the file as CopyStyleNumberFormatDemo.xlsx.
    class CopyStyleNumberFormatDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ---------- Source cell with currency format ----------
                // Put a numeric value in A1
                Cell srcCell = sheet.Cells["A1"];
                srcCell.PutValue(1234.56);

                // Create a style and set a built‑in currency number format (ID 5)
                Style srcStyle = workbook.CreateStyle();
                srcStyle.Number = 5; // "$#,##0_);($#,##0)" – displays currency symbol
                srcCell.SetStyle(srcStyle);

                // ---------- Destination range ----------
                // Define source and destination ranges
                CellsRange srcRange = sheet.Cells.CreateRange("A1");
                CellsRange destRange = sheet.Cells.CreateRange("B2");

                // Copy the style (including number format) from source to destination
                destRange.CopyStyle(srcRange);

                // Put a value in the destination cell to see the inherited format
                sheet.Cells["B2"].PutValue(9876.54);

                // Save the workbook
                workbook.Save("CopyStyleNumberFormatDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            CopyStyleNumberFormatDemo.Run();
        }
    }
}
