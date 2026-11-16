// Title: Insert columns before a header and refresh freeze panes with Aspose.Cells for .NET (C#)
// Common Searches: C# insert multiple columns before column C using Aspose.Cells and keep freeze panes | how to reset freeze panes after adding columns in an Excel file with Aspose.Cells | Aspose.Cells FreezePanes overload to recalculate frozen column after column insertion | insert three columns at index 2 and reapply freeze panes in a workbook using C# | preserve frozen rows and columns when inserting columns with Aspose.Cells .NET
// Tags: insert columns before specific index Aspose.Cells | reapply freeze panes after column insertion Aspose.Cells | Aspose.Cells FreezePanes overload usage | C# modify worksheet layout Aspose.Cells | adjust frozen column index Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, inserts three columns before column C (zero‑based index 2) on the first worksheet, resets any existing freeze panes using the four‑parameter FreezePanes overload, and saves the updated file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists.
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Insert three columns before column C (zero‑based index 2).
                int insertAtColumnIndex = 2;
                int numberOfColumnsToInsert = 3;
                sheet.Cells.InsertColumns(insertAtColumnIndex, numberOfColumnsToInsert);

                // Reset freeze panes (if any) after column insertion.
                // The 4‑parameter overload sets the freeze row, column and the pane's top‑left cell.
                sheet.FreezePanes(0, 0, 0, 0);

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
