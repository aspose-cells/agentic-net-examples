// Title: C# – Auto‑fit a Row After Updating a Cell with Aspose.Cells for .NET
// Description: Loads an existing workbook, changes the text of cell B2 on the first worksheet, automatically adjusts the height of row 2 to fit the new content, and saves the result as a new XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells AutoFitRow C# | update cell value Aspose.Cells | adjust row height programmatically | worksheet.AutoFitRow example | .NET spreadsheet row auto‑fit
// Common Searches: Aspose.Cells how to auto‑fit a row after editing a cell | C# auto‑fit specific row Aspose.Cells | adjust row height based on new cell text .NET | AutoFitRow method usage Aspose.Cells
// Developer Intent: Change a cell’s content and have the corresponding row automatically resize to display the full text.
// Use Cases: Resize a header row after extending its label to avoid truncation. | Ensure report rows keep proper height after bulk text updates. | Maintain consistent layout in generated workbooks by auto‑fitting rows post‑modification.
// AI Prompts: Create C# code that updates several cells and calls AutoFitRow for each affected row using Aspose.Cells. | Show how to auto‑fit a range of rows after batch editing cells in a worksheet with Aspose.Cells for .NET. | Provide an example of using AutoFitRow with merged cells and wrapped text in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitRowExample
{
    // Loads an existing workbook, changes the text of cell B2 on the first worksheet, automatically adjusts the height of row 2 to fit the new content, and saves the result as a new XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Modify the value of cell B2 (row index 1, column index 1)
            worksheet.Cells["B2"].PutValue("Updated cell value with a longer text to demonstrate auto‑fit.");

            // Auto‑fit the row that contains the modified cell (row index 1)
            worksheet.AutoFitRow(1);

            // Save the updated workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
