// Title: Open an XLSX workbook from a file using Aspose.Cells for .NET (C#)
// Description: This example shows how to instantiate a Workbook from a local XLSX file, access the first worksheet, and print its name along with the count of populated rows and columns using Aspose.Cells APIs.
// Keywords: Aspose.Cells C# load workbook | open Excel file .NET | Workbook constructor file path | first worksheet name Aspose | used rows columns count | read XLSX metadata | Aspose.Cells file I/O
// Common Searches: Aspose.Cells how to open XLSX file in C# | C# get first sheet name with Aspose.Cells | retrieve used row count Aspose.Cells | display worksheet dimensions Aspose.Cells | load workbook from disk Aspose.Cells example
// Developer Intent: Open an Excel file and read basic information about its first worksheet.
// Use Cases: Validate incoming Excel files by checking sheet names and data extents before processing. | Extract workbook metadata to drive dynamic UI elements or reporting templates. | Quickly assess file structure (row/column usage) for conditional logic in data pipelines.
// AI Prompts: Generate C# code that uses Aspose.Cells to open an XLSX file and output the first worksheet's name, used rows, and used columns. | Create a reusable method that accepts a file path, loads the workbook with Aspose.Cells, and returns an object containing sheet count, first sheet name, and data dimensions. | Explain best practices for handling exceptions and file‑access errors when loading a workbook with Aspose.Cells in C#.

using System;
using Aspose.Cells;

// This example shows how to instantiate a Workbook from a local XLSX file, access the first worksheet, and print its name along with the count of populated rows and columns using Aspose.Cells APIs.
class LoadWorkbookDemo
{
    static void Main()
    {
        // Path to the XLSX file on disk
        string filePath = "example.xlsx";

        // Load the workbook using the Workbook(string) constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the loaded workbook
        Worksheet firstSheet = workbook.Worksheets[0];

        // Display basic information about the loaded workbook
        Console.WriteLine("Workbook loaded successfully.");
        Console.WriteLine("First worksheet name: " + firstSheet.Name);
        Console.WriteLine("Used rows: " + (firstSheet.Cells.MaxDataRow + 1));
        Console.WriteLine("Used columns: " + (firstSheet.Cells.MaxDataColumn + 1));
    }
}
