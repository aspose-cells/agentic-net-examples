// Title: C# Load Excel workbook from file path and retrieve first worksheet using Aspose.Cells
// Description: Demonstrates how to verify a file’s existence, create a Workbook from a given path, access the first worksheet (index 0), and print its name to the console, with robust exception handling.
// Keywords: Aspose.Cells | C# load workbook | open Excel file | first worksheet | file existence check | Workbook constructor | console output
// Common Searches: Aspose.Cells open Excel file C# | Get first sheet name Aspose.Cells | Check if Excel file exists before loading C# | Load workbook and read worksheet name .NET
// Developer Intent: Open an Excel file by path and obtain the name of its first worksheet.
// Use Cases: Log the first sheet name when processing uploaded Excel files. | Use the first worksheet as a data source for import routines. | Validate template structure before generating reports.
// AI Prompts: Generate C# code that safely opens an Excel file with Aspose.Cells, checks for existence, and prints the first worksheet name. | Create a robust error‑handling pattern for loading a workbook when the file may be missing or corrupted. | Show how to load a workbook, access the first worksheet, and append a row of data using Aspose.Cells in C#.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to verify a file’s existence, create a Workbook from a given path, access the first worksheet (index 0), and print its name to the console, with robust exception handling.
public class LoadWorkbookExample
{
    public static void Run(string filePath)
    {
        try
        {
            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook from the specified file path
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (zero‑based index)
            Worksheet firstWorksheet = workbook.Worksheets[0];

            // Display the name of the first worksheet
            Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Determine the workbook path: use argument if provided, otherwise a default placeholder
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        // Execute the example with safety checks
        LoadWorkbookExample.Run(filePath);
    }
}
