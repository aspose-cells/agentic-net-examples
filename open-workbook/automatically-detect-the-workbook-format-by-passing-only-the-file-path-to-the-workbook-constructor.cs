// Title: Load any Excel or CSV workbook by passing only the file path to Aspose.Cells Workbook constructor in C#
// AI Prompts: Generate a C# example that creates a new Aspose.Cells Workbook using just a file path, allowing the library to automatically detect XLS, XLSX, CSV, or other supported formats, then prints the total number of worksheets. | Show how to open a workbook with an unknown extension in C# using Aspose.Cells, access its Worksheets collection, and demonstrate that no format argument is required.
// Common Searches: Aspose.Cells C# open workbook without specifying file type | How to let Aspose.Cells detect XLSX or CSV format from file path in .NET | C# load Excel file of unknown extension using Aspose.Cells Workbook constructor | Retrieve worksheet count after auto-detecting workbook format with Aspose.Cells
// Tags: auto-detect workbook format Aspose.Cells | load workbook from file path C# | open XLSX or CSV without format parameter Aspose.Cells | instantiate Workbook with path detection | retrieve worksheet count Aspose.Cells

using System;
using Aspose.Cells;

// // Demonstrates using the Aspose.Cells Workbook constructor to open a workbook from a file path, letting the library automatically determine the file format (XLS, XLSX, CSV, etc.) and then outputting the number of worksheets loaded.
class Program
{
    static void Main()
    {
        // Specify the path to the workbook file (XLS, XLSX, CSV, etc.).
        string filePath = @"C:\Data\SampleWorkbook.xlsx";

        // The Workbook constructor automatically detects the file format.
        Workbook workbook = new Workbook(filePath);

        // Example usage: output the number of worksheets loaded.
        Console.WriteLine($"Worksheets loaded: {workbook.Worksheets.Count}");
    }
}
