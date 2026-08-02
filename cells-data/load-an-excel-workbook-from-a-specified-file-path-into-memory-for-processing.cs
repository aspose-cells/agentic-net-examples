// Title: C# – Load an Excel Workbook from a File Path using Aspose.Cells
// Description: Demonstrates how to instantiate an Aspose.Cells Workbook in .NET by passing a full file path to the Workbook constructor. The example shows loading the workbook into memory, accessing the first worksheet, and reading cell A1 without locking the source file.
// Keywords: Aspose.Cells load workbook C# | open Excel file .NET | Workbook constructor file path | read Excel cell A1 Aspose | load Excel into memory C#
// Common Searches: How to open an .xlsx file with Aspose.Cells in C# | Load workbook from specific path Aspose.Cells .NET | Get worksheet name after loading Excel file Aspose | Read cell value after opening workbook Aspose.Cells
// Developer Intent: Create a Workbook object from a local Excel file for further manipulation.
// Use Cases: Read data from a known .xlsx file on disk. | Iterate through worksheets after loading a workbook. | Apply formatting, formulas, or calculations to a workbook that has just been opened.
// AI Prompts: Show C# code that loads an Excel workbook from a path with Aspose.Cells and catches file‑not‑found errors. | Provide an example that opens a workbook, prints the first worksheet name, and returns the value of cell A1. | Explain how to load a workbook into memory without keeping the source file locked using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate an Aspose.Cells Workbook in .NET by passing a full file path to the Workbook constructor. The example shows loading the workbook into memory, accessing the first worksheet, and reading cell A1 without locking the source file.
public class WorkbookLoader
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <returns>Instance of <see cref="Workbook"/> representing the loaded file.</returns>
    public Workbook LoadWorkbook(string filePath)
    {
        // Use the Workbook(string) constructor which opens the file directly.
        Workbook workbook = new Workbook(filePath);
        return workbook;
    }
}

// Example usage
class Program
{
    static void Main()
    {
        string excelPath = @"C:\Data\Sample.xlsx";

        WorkbookLoader loader = new WorkbookLoader();
        Workbook wb = loader.LoadWorkbook(excelPath);

        // Access the first worksheet to demonstrate that the workbook is loaded.
        Worksheet sheet = wb.Worksheets[0];
        Console.WriteLine($"Loaded worksheet name: {sheet.Name}");
        Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");
    }
}
