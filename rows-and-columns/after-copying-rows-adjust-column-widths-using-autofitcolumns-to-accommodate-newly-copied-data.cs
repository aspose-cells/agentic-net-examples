// Title: Copy rows and auto‑fit columns with Aspose.Cells for .NET (C# example)
// Description: C# sample that copies selected rows from a source worksheet to a destination worksheet using Aspose.Cells CopyRows, then automatically resizes all columns with AutoFitColumns before saving the workbook.
// Keywords: Aspose.Cells | C# | CopyRows | AutoFitColumns | adjust column width | copy rows between worksheets | Excel automation | Aspose.Cells example | copy rows C#
// Common Searches: Aspose.Cells copy rows C# | AutoFitColumns after copying rows Aspose.Cells | how to auto size columns after CopyRows .NET | copy rows between worksheets Aspose.Cells example
// Developer Intent: Copy specific rows from one worksheet to another and automatically resize the columns to fit the newly copied data using Aspose.Cells for .NET.
// Use Cases: Replicate header rows in a new report workbook while keeping column widths optimal. | Extract a subset of data from a template sheet and paste it into a generated sheet with readable column sizes. | Combine rows from multiple source workbooks into a single file and apply AutoFitColumns for consistent formatting.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy rows from one worksheet to another and then calls AutoFitColumns on the target sheet. | Explain how AutoFitColumns works after a CopyRows operation and how to limit the auto‑fit to a specific column range. | Provide an Aspose.Cells .NET example that copies rows, auto‑fits columns, and saves the workbook with a custom file name.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // C# sample that copies selected rows from a source worksheet to a destination worksheet using Aspose.Cells CopyRows, then automatically resizes all columns with AutoFitColumns before saving the workbook.
    public class CopyRowsAndAutoFitColumnsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a source workbook and populate it with sample data
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Fill some rows with data that have varying column widths
            sourceSheet.Cells["A1"].PutValue("Short");
            sourceSheet.Cells["B1"].PutValue("This is a much longer text that should expand column B");
            sourceSheet.Cells["C1"].PutValue(12345);

            sourceSheet.Cells["A2"].PutValue("Another short");
            sourceSheet.Cells["B2"].PutValue("Medium length text");
            sourceSheet.Cells["C2"].PutValue(67890);

            // Create a destination workbook (empty)
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "Destination";

            // Copy the first two rows from source to destination starting at row index 0
            destSheet.Cells.CopyRows(sourceSheet.Cells, 0, 0, 2);

            // Auto-fit columns in the destination worksheet
            destSheet.AutoFitColumns();

            // Save the destination workbook
            string outputPath = "CopyRowsAndAutoFitColumnsResult.xlsx";
            destWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
