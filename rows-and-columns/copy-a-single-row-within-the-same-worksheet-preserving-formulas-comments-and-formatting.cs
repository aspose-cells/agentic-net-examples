// Title: Copy a Single Row with Formulas, Comments, and Formatting Using Aspose.Cells for .NET
// Description: Demonstrates how to duplicate a row in the same worksheet with Aspose.Cells for .NET while retaining cell values, formulas, styles, and comments, then saves the workbook as an XLSX file.
// Keywords: Aspose.Cells CopyRow | copy Excel row C# | preserve formulas Aspose.Cells | copy row with comments | duplicate worksheet row .NET | Excel row copy formatting | Aspose.Cells row copy example
// Common Searches: Aspose.Cells copy row preserving formulas | C# copy Excel row with comments | How to duplicate a worksheet row using Aspose.Cells | CopyRow method example Aspose.Cells .NET | Copy row with formatting Aspose.Cells
// Developer Intent: Duplicate a row within the same worksheet while keeping all cell content, formulas, styles, and comments unchanged.
// Use Cases: Create a reusable header row for data entry templates. | Apply the same calculation logic to additional rows without rewriting formulas. | Clone a documented row with notes for training or audit trails.
// AI Prompts: Write C# code with Aspose.Cells to copy row 3 to row 12, preserving formulas, comments, and cell styles. | Explain the parameters of Cells.CopyRow and how it handles merged cells and comments in Aspose.Cells for .NET. | Suggest robust error‑handling patterns when copying rows that contain formulas and comments using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to duplicate a row in the same worksheet with Aspose.Cells for .NET while retaining cell values, formulas, styles, and comments, then saves the workbook as an XLSX file.
    public class CopySingleRowDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // Use new Workbook("input.xlsx") to load an existing file

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate the source row (row 0) with data, a formula, and a comment
                cells["A1"].PutValue("Sample Text");
                cells["B1"].PutValue(10);
                cells["C1"].Formula = "=B1*2";

                // Add a comment to cell A1
                int commentIdx = sheet.Comments.Add("A1");
                Comment comment = sheet.Comments[commentIdx];
                comment.Note = "This is a comment on the source row";

                // Define source and destination row indices (zero‑based)
                int sourceRowIndex = 0;      // Row 1 in Excel
                int destinationRowIndex = 4; // Row 5 in Excel

                // Copy the entire row preserving values, formulas, formatting, and comments
                cells.CopyRow(cells, sourceRowIndex, destinationRowIndex);

                // Save the workbook with the copied row
                string outputPath = "CopyRowResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CopySingleRowDemo.Run();
        }
    }
}
