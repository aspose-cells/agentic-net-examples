// Title: Auto‑fit a worksheet row after updating a cell value using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing XLSX file, changes the text in cell B2, calls Worksheet.AutoFitRow for that row, and saves the workbook. | Generate a snippet that demonstrates modifying a cell value and automatically adjusting the row height with Aspose.Cells in a .NET console application. | Provide a step‑by‑step example of loading a workbook, updating a cell, invoking AutoFitRow on the affected row, and persisting the changes.
// Common Searches: aspocells c# autofitrow after cell edit example | how to resize a specific row when cell content changes using Aspose.Cells | c# load workbook modify cell and auto‑fit row height | worksheet.AutoFitRow usage for single row in .NET Excel processing
// Tags: auto-fit row Aspose.Cells C# | modify cell value worksheet Aspose.Cells | Worksheet.AutoFitRow method example | adjust row height after cell update .NET | load and save Excel workbook Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitRowExample
{
    // The sample loads an existing XLSX workbook, updates cell B2 with longer text, automatically fits the second row to the new content using Worksheet.AutoFitRow, and saves the modified file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify a cell value (e.g., B2). Row and column indexes are zero‑based.
            worksheet.Cells["B2"].PutValue("Updated cell content with a longer text to demonstrate auto‑fit.");

            // Auto‑fit the row that contains the modified cell (row index 1 corresponds to row 2 in Excel)
            worksheet.AutoFitRow(1);

            // Save the updated workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
