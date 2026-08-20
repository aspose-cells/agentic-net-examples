// Title: C# Console App: Freeze Panes in Excel with Aspose.Cells Using Command‑Line Row/Column Indices
// Description: A lightweight console utility that reads a row index and a column index from the command line, validates the inputs, creates a new workbook, applies Worksheet.FreezePanes at the specified position, saves the file as FreezePanesOutput.xlsx, and prints status messages.
// Keywords: Aspose.Cells | C# | FreezePanes | command line | row index | column index | Excel automation | console application | programmatic freeze panes | Aspose.Cells example
// Common Searches: Aspose.Cells freeze panes from command line | C# freeze rows and columns in Excel programmatically | how to use Worksheet.FreezePanes with arguments | console app to lock Excel headers Aspose | run FreezePanes demo with parameters
// Developer Intent: Implement a command‑line driven C# program that freezes specific rows and columns in an Excel worksheet using Aspose.Cells.
// Use Cases: Create a reusable CLI tool for end‑users to set freeze panes before distributing reports. | Integrate freeze‑pane logic into automated build pipelines that generate Excel dashboards. | Provide a quick‑start example for developers learning how to control worksheet view settings via code.
// AI Prompts: Write C# code that reads row and column numbers from the command line and calls Worksheet.FreezePanes with Aspose.Cells. | Add comprehensive error handling to the FreezePanes console app, including out‑of‑range checks and custom messages. | Show how to modify the example to freeze only the top rows while keeping columns scrollable.

using System;
using Aspose.Cells;

namespace FreezePanesDemo
{
    // A lightweight console utility that reads a row index and a column index from the command line, validates the inputs, creates a new workbook, applies Worksheet.FreezePanes at the specified position, saves the file as FreezePanesOutput.xlsx, and prints status messages.
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two command‑line arguments: row index and column index for freezing.
            // Example usage: FreezePanesDemo.exe 3 3
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: <exe> <rowIndex> <columnIndex>");
                return;
            }

            // Parse the row and column indices.
            if (!int.TryParse(args[0], out int rowIndex) || rowIndex < 0)
            {
                Console.WriteLine("Invalid row index.");
                return;
            }

            if (!int.TryParse(args[1], out int columnIndex) || columnIndex < 0)
            {
                Console.WriteLine("Invalid column index.");
                return;
            }

            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze panes at the specified cell.
            // The last two parameters define how many rows and columns are frozen,
            // which in this simple scenario are the same as the freeze position.
            worksheet.FreezePanes(rowIndex, columnIndex, rowIndex, columnIndex);

            // Save the workbook to a file.
            string outputPath = "FreezePanesOutput.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Freeze panes applied at row {rowIndex}, column {columnIndex}.");
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
