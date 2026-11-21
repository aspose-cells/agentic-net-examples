// Title: Load an Excel workbook from a file into memory with Aspose.Cells for .NET (C#)
// Common Searches: how to read an .xlsx file into a Workbook object using Aspose.Cells in C# | Aspose.Cells C# code to open an Excel workbook from a file path for further processing | example of loading an Excel workbook into memory with Aspose.Cells .NET library | C# Aspose.Cells load workbook from local file and access first worksheet | using Aspose.Cells to import an existing .xlsx into a Workbook instance in .NET
// Tags: Aspose.Cells load workbook from file | C# read .xlsx into memory using Aspose.Cells | initialize Workbook with file name in .NET | open Excel workbook for further processing in C# | Aspose.Cells workbook loading pattern | load Excel file into Workbook instance C#

using Aspose.Cells;
using System;

// The code creates a Workbook object by loading the 'input.xlsx' file from disk, then accesses the first worksheet, making the workbook ready for additional manipulation.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook from the file into memory
        Workbook workbook = new Workbook(filePath);

        // The workbook is now ready for further manipulation
        // Example: access the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];

        // Additional processing can be performed here
    }
}
