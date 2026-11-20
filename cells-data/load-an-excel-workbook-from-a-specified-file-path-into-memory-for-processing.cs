// Title: Load an Excel workbook from a file path into memory with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load workbook from local file path | Read Excel file into memory using Aspose.Cells Workbook constructor | Get first worksheet name after opening .xlsx with Aspose.Cells in .NET | How to open an .xlsx file in C# without using Interop | Initialize Aspose.Cells Workbook with absolute path in C#
// Tags: open workbook via file path Aspose.Cells | create Workbook instance C# | read .xlsx into memory Aspose.Cells | retrieve first worksheet name Aspose.Cells | Aspose.Cells workbook loading example

using System;
using Aspose.Cells;

// The sample defines a full file path to an .xlsx workbook, creates a Workbook object with that path to load the file into memory, and then accesses the first worksheet's Name property, outputting it to the console.
class ExcelProcessor
{
    static void Main()
    {
        // Specify the full path to the Excel file you want to load.
        string filePath = @"C:\Data\SampleWorkbook.xlsx";

        // Load the workbook into memory.
        // The Workbook constructor reads the file and creates an in‑memory representation.
        Workbook workbook = new Workbook(filePath);

        // At this point you can work with the workbook (e.g., read sheets, cells, etc.).
        // Example: Get the name of the first worksheet.
        string firstSheetName = workbook.Worksheets[0].Name;
        Console.WriteLine($"First worksheet name: {firstSheetName}");
    }
}
