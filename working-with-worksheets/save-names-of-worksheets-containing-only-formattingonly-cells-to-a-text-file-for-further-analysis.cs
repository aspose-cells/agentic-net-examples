// Title: Write worksheet names that contain only formatting (no data) to a text file with Aspose.Cells for .NET
// AI Prompts: Generate C# code that scans an Excel workbook using Aspose.Cells, identifies worksheets where MaxDataRow equals -1, and writes each worksheet name to a new line in a .txt file. | Create a reusable C# method with Aspose.Cells that returns a list of sheet names that have only style information and saves the list to a user‑specified text file path. | Provide a C# snippet that records both the index and name of formatting‑only worksheets into a CSV file using Aspose.Cells.
// Common Searches: Aspose.Cells C# find worksheets with only styles and export their names | how to list Excel sheets that have no data rows but contain formatting using Aspose.Cells | save names of formatting‑only worksheets to a text file in .NET | detect empty worksheets with formatting in Aspose.Cells and write results to file
// Tags: identify formatting‑only worksheets Aspose.Cells | export worksheet names to txt C# | use MaxDataRow to detect empty sheets | list Excel sheets with styles but no values | save formatting‑only sheet list Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through each worksheet, uses MaxDataRow to detect sheets that contain only formatting (no data), collects their names, and writes the names line‑by‑line to a text file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // List to hold names of worksheets that contain only formatting (no data)
        List<string> formattingOnlySheets = new List<string>();

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow returns -1 when the sheet has no cells with values
            // Cells that contain only style information are not counted as data
            if (sheet.Cells.MaxDataRow == -1)
            {
                formattingOnlySheets.Add(sheet.Name);
            }
        }

        // Write the worksheet names to a text file (one name per line)
        // Replace with your desired output path
        File.WriteAllLines("FormattingOnlySheets.txt", formattingOnlySheets);
    }
}
