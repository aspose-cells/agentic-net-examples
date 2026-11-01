// Title: C# Aspose.Cells example: Load an Excel workbook from a file path into memory
// AI Prompts: Write C# code that uses Aspose.Cells to open a .xlsx file from a given absolute path and prints the value of cell A1. | Show how to load an Excel workbook into a Workbook object using a FileStream with Aspose.Cells in C#. | Provide a C# snippet that loads a workbook, accesses the first worksheet, and catches FileNotFoundException when the file path is invalid.
// Common Searches: how to open an existing Excel file with Aspose.Cells in C# | Aspose.Cells C# load workbook from specific path and read cell value | C# example for reading A1 from .xlsx using Aspose.Cells library | using Aspose.Cells to load workbook with FileStream instead of direct path | Aspose.Cells C# error handling when loading workbook from file
// Tags: load workbook from file path Aspose.Cells C# | read cell A1 Aspose.Cells | initialize Workbook with FileStream Aspose.Cells | handle missing Excel file Aspose.Cells | access first worksheet Aspose.Cells API

using Aspose.Cells;
using System;

// // Demonstrates loading an Excel workbook from a specified file path using Aspose.Cells in C#, accessing the first worksheet, reading the value of cell A1, and outputting it to the console.
class Program
{
    static void Main(string[] args)
    {
        // Path to the Excel file to be loaded
        string filePath = @"C:\Path\To\Your\File.xlsx";

        // Load the workbook from the specified file into memory
        Workbook workbook = new Workbook(filePath);

        // Example processing: access the first worksheet and read a cell value
        Worksheet worksheet = workbook.Worksheets[0];
        string cellValue = worksheet.Cells["A1"].StringValue;
        Console.WriteLine($"Value in A1: {cellValue}");
    }
}
