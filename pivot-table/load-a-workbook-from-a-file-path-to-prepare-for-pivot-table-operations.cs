// Title: Load an Excel workbook from a file path using Aspose.Cells for .NET – ready for pivot table work
// Description: Demonstrates how to instantiate a Workbook object by passing the Excel file path to Aspose.Cells for .NET. The file is read into memory, its worksheets become accessible, and the workbook can immediately be used for creating or modifying pivot tables.
// Keywords: Aspose.Cells | C# | load workbook | Excel file path | pivot table preparation | Workbook constructor | .NET Excel API | open Excel file
// Common Searches: Aspose.Cells load workbook from file path C# | how to open an existing Excel file with Aspose.Cells | initialize workbook for pivot tables Aspose.Cells .NET | read Excel file into Aspose.Cells Workbook object | C# Aspose.Cells open workbook for pivot table creation
// Developer Intent: Open an existing Excel file and create a Workbook instance so that pivot tables can be added or edited programmatically.
// Use Cases: Load a workbook and verify the number of worksheets before building a pivot table. | Open a file, scan worksheets to locate the data range that will serve as the pivot cache. | Initialize a workbook from disk and pass it to Aspose.Cells pivot‑table APIs to generate a new report.
// AI Prompts: Generate C# code that adds a pivot table to the first worksheet of a workbook loaded from a file path using Aspose.Cells. | Explain how to catch and handle exceptions when opening an Excel file with Aspose.Cells. | Show how to verify that the source file exists before creating the Workbook object in C#.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate a Workbook object by passing the Excel file path to Aspose.Cells for .NET. The file is read into memory, its worksheets become accessible, and the workbook can immediately be used for creating or modifying pivot tables.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "example.xlsx";

        // Load the workbook from the specified file path
        Workbook workbook = new Workbook(filePath);

        // The workbook is now ready for pivot table operations
        // Example: display the number of worksheets loaded
        Console.WriteLine($"Worksheets loaded: {workbook.Worksheets.Count}");
    }
}
