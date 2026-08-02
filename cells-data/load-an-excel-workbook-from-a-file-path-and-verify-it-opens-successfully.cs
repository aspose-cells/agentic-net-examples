// Title: Load and Verify an Excel Workbook from a File Path with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open an existing Excel file by passing its path to the Aspose.Cells Workbook constructor, confirm the load by reading the worksheet count, and handle any errors with a try‑catch block.
// Keywords: Aspose.Cells load workbook C# | open Excel file Aspose.Cells | verify workbook opened Aspose | worksheet count Aspose.Cells | exception handling Aspose.Cells
// Common Searches: How to open an Excel file with Aspose.Cells in C# | Check if workbook loaded successfully using Aspose.Cells | Get number of worksheets after loading Excel with Aspose | Aspose.Cells C# error handling when opening a file
// Developer Intent: Open an Excel file via its file system path, ensure the workbook is loaded, and retrieve basic validation data such as worksheet count.
// Use Cases: Validate user‑uploaded Excel files before processing data. | Log workbook structure (sheet count) for audit or debugging purposes. | Gracefully report file‑access or format errors to prevent downstream failures.
// AI Prompts: Write C# code that loads an Excel workbook from a given path using Aspose.Cells, returns the worksheet count, and includes robust exception handling. | Show how to add a pre‑load file‑existence check to the Aspose.Cells workbook loading example. | Suggest best practices for logging and error reporting when opening Excel files with Aspose.Cells in a .NET application.

using System;
using Aspose.Cells;

// Demonstrates how to open an existing Excel file by passing its path to the Aspose.Cells Workbook constructor, confirm the load by reading the worksheet count, and handle any errors with a try‑catch block.
class LoadWorkbookDemo
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "example.xlsx";

        try
        {
            // Load the workbook using the string constructor (opens the file)
            Workbook workbook = new Workbook(filePath);

            // Verify that the workbook opened by checking the number of worksheets
            Console.WriteLine($"Workbook loaded successfully. Worksheet count: {workbook.Worksheets.Count}");
        }
        catch (Exception ex)
        {
            // If an exception occurs, the load failed
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
        }
    }
}
