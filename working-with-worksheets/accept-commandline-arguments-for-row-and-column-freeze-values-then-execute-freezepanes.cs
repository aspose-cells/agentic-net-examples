// Title: Freeze rows and columns from the command line using Aspose.Cells for .NET (C#)
// Description: A console utility that reads optional zero‑based row and column indices from the command line, defaults to row 1/column 1 when omitted or invalid, creates a workbook, applies Worksheet.FreezePanes at the specified position, and saves the file as FreezePanesOutput.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | command line Excel freeze rows | freeze columns programmatically .NET | Worksheet.FreezePanes example | Excel pane freezing console app
// Common Searches: Aspose.Cells freeze panes from command line | C# set freeze row and column in Excel workbook | default freeze position Aspose.Cells | parse command line arguments for FreezePanes | how to lock top rows and left columns using Aspose.Cells
// Developer Intent: Read row and column indices supplied via command‑line arguments and apply them to Worksheet.FreezePanes.
// Use Cases: Create a reusable CLI tool that lets end‑users define which header rows and side columns stay visible in generated reports. | Automate workbook preparation where the freeze location varies per execution without changing source code. | Provide a fallback freeze at the first row/column when no valid arguments are passed, ensuring a consistent view for all users.
// AI Prompts: Generate C# code that accepts row, column, and optional output file name as command‑line parameters and uses Aspose.Cells to freeze panes and save the workbook. | Explain the four parameters of Worksheet.FreezePanes and how they control the number of frozen rows and columns in an Excel sheet.

using System;
using Aspose.Cells;

namespace FreezePanesDemo
{
    // A console utility that reads optional zero‑based row and column indices from the command line, defaults to row 1/column 1 when omitted or invalid, creates a workbook, applies Worksheet.FreezePanes at the specified position, and saves the file as FreezePanesOutput.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Default freeze position if arguments are not provided
            int freezeRow = 1;    // Row index (0‑based)
            int freezeColumn = 1; // Column index (0‑based)

            // Parse command‑line arguments: first = row, second = column
            if (args.Length >= 2)
            {
                // Try to convert arguments to integers; fall back to defaults on failure
                if (!int.TryParse(args[0], out freezeRow))
                    freezeRow = 1;
                if (!int.TryParse(args[1], out freezeColumn))
                    freezeColumn = 1;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze panes at the specified cell.
            // The last two parameters define how many rows/columns are frozen,
            // which we set equal to the freeze position to freeze the top/left area.
            worksheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

            // Save the workbook
            workbook.Save("FreezePanesOutput.xlsx");
        }
    }
}
