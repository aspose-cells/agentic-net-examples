// Title: C# – Merge B2:D2 into one cell and center text horizontally with Aspose.Cells for .NET
// Description: A concise C# example that creates a workbook, merges the range B2:D2 on the first worksheet, inserts "Merged and Centered", applies horizontal center alignment, and saves the file as MergedCells_B2_D2.xlsx using the Aspose.Cells for .NET API.
// Keywords: Aspose.Cells merge cells C# | merge B2 D2 Aspose.Cells | horizontal center alignment Aspose.Cells | C# Excel cell merging | Aspose.Cells style alignment | Aspose.Cells Workbook example | .NET Excel merge cells | Aspose.Cells API header merge
// Common Searches: Aspose.Cells merge cells C# | How to merge B2:D2 with Aspose.Cells | Center text in merged cells using Aspose.Cells .NET | C# code to merge Excel cells and align horizontally | Aspose.Cells set cell style horizontal alignment
// Developer Intent: Merge cells B2 through D2 into a single cell and horizontally center its content using Aspose.Cells for .NET.
// Use Cases: Create a spanning header row for a generated financial report. | Design a centered title cell in a reusable spreadsheet template. | Combine label cells for invoices, receipts, or dashboards where alignment matters.
// AI Prompts: Generate C# code with Aspose.Cells to merge A1:C1, make the text bold, and apply center alignment. | Show how to merge a cell range and set vertical alignment using Aspose.Cells for .NET. | Explain how to merge cells while preserving existing data in the source range with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A concise C# example that creates a workbook, merges the range B2:D2 on the first worksheet, inserts "Merged and Centered", applies horizontal center alignment, and saves the file as MergedCells_B2_D2.xlsx using the Aspose.Cells for .NET API.
    class MergeCellsExample
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Merge cells B2:D2 (row 1, column 1, 1 row, 3 columns)
            cells.Merge(1, 1, 1, 3);

            // Put a value into the merged cell (upper‑left cell B2)
            cells[1, 1].PutValue("Merged and Centered");

            // Center the text horizontally
            Style style = cells[1, 1].GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            cells[1, 1].SetStyle(style);

            // Define output file path
            string outputPath = "MergedCells_B2_D2.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}
