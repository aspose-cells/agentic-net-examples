// Title: How to freeze the first row and column in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new workbook with Aspose.Cells, freezes the topmost row and leftmost column, and saves the file. | Demonstrate using the Worksheet.FreezePanes method to lock header rows and columns in an Excel file with Aspose.Cells. | Provide a step‑by‑step example that sets FreezePanes(1,1,1,1) on the first worksheet and outputs the workbook to a .xlsx file.
// Common Searches: Aspose.Cells C# freeze top row and left column example | How to keep Excel header row visible while scrolling using Aspose.Cells .NET | Worksheet.FreezePanes parameters explanation for freezing first row and column | C# code to freeze panes in an Excel file with Aspose.Cells | Saving a workbook after applying FreezePanes in Aspose.Cells
// Tags: Aspose.Cells freeze panes C# | Worksheet.FreezePanes first row column | Excel header freeze Aspose.Cells .NET | C# save workbook after FreezePanes | Aspose.Cells set frozen pane parameters

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a new workbook, accesses the first worksheet, applies Worksheet.FreezePanes(1,1,1,1) to lock the top row and left column, saves the workbook as 'FrozenPaneExample.xlsx', and outputs the file path.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Freeze the first row and first column.
                // Parameters: row, column, totalRows, totalColumns
                // row/column specify the top‑left cell of the scrollable area (cell after the frozen area).
                // totalRows/totalColumns specify how many rows/columns to freeze.
                worksheet.FreezePanes(1, 1, 1, 1);

                // Define output file path
                string outputPath = "FrozenPaneExample.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
