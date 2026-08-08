// Title: Import File Path Array and Create Adjacent Hyperlinks with Aspose.Cells (C#)
// Description: Shows how to import a string array of file paths into column A, add a hyperlink in column B that opens each file, display only the file name, and save the workbook as FileLinks.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# import string array | Excel hyperlink | CellsHelper.CellIndexToName | Hyperlinks.Add | .NET file links | generate Excel index | file path list
// Common Searches: Aspose.Cells import string array C# | Add hyperlink to Excel cell using Aspose.Cells | Create file links in Excel with Aspose.Cells .NET | How to generate clickable file paths in Excel programmatically | Aspose.Cells CellsHelper example
// Developer Intent: Create an Excel worksheet that lists file paths and provides a clickable link for each entry.
// Use Cases: Document catalog where each row shows a path and a direct‑open link. | Server‑based PDF index sheet displaying only file names as hyperlinks. | Automated report that converts a dynamic list of locations into clickable Excel links.
// AI Prompts: Write C# code with Aspose.Cells to import a string array of file paths into column A and add hyperlinks in column B that display the file names. | Show how to use CellsHelper.CellIndexToName and Hyperlinks.Add to generate file hyperlinks for each array element. | Provide an example that saves the workbook after adding hyperlinks to file paths, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to import a string array of file paths into column A, add a hyperlink in column B that opens each file, display only the file name, and save the workbook as FileLinks.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Array of file paths to be imported
        string[] filePaths = new string[]
        {
            @"C:\Docs\file1.pdf",
            @"C:\Docs\file2.pdf",
            @"C:\Docs\file3.pdf"
        };

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Import the file paths vertically starting at cell A1
        // Parameters: (stringArray, firstRow, firstColumn, isVertical)
        sheet.Cells.ImportArray(filePaths, 0, 0, true);

        // Add a hyperlink in the adjacent column (B) for each file path
        for (int i = 0; i < filePaths.Length; i++)
        {
            // Convert row/column indices to an Excel cell name (e.g., B1, B2, ...)
            string hyperlinkCell = CellsHelper.CellIndexToName(i, 1); // column index 1 = B

            // Add the hyperlink pointing to the file path
            // Parameters: (cellName, totalRows, totalColumns, address)
            sheet.Hyperlinks.Add(hyperlinkCell, 1, 1, filePaths[i]);

            // Optionally set the display text of the hyperlink to the file name
            sheet.Cells[hyperlinkCell].PutValue(System.IO.Path.GetFileName(filePaths[i]));
        }

        // Save the workbook to a file
        workbook.Save("FileLinks.xlsx");
    }
}
