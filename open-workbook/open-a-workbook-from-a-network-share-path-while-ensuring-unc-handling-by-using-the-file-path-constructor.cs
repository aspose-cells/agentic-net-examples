// Title: Load an Excel workbook from a UNC network share path using Aspose.Cells for .NET
// AI Prompts: Verify that a UNC file exists with File.Exists before creating a Workbook instance via Aspose.Cells. | Wrap the Workbook constructor in a try‑catch block to gracefully handle I/O or format errors when opening a file from a network share. | After the workbook is loaded, read the Name property of the first Worksheet and write it to the console.
// Common Searches: how to open an Excel file from a UNC share using Aspose.Cells in C# | Aspose.Cells check file existence before loading workbook from network location | catch exceptions when loading workbook from a network path with Aspose.Cells | read first worksheet name after loading workbook from UNC path in .NET
// Tags: Aspose.Cells network share file loading | C# UNC path existence verification | Workbook constructor error handling | Extract first worksheet name with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates checking for the presence of an Excel file on a UNC network share, loading it with Aspose.Cells' Workbook constructor, printing the name of the first worksheet, and handling any errors that may occur during the process.
class Program
{
    static void Main()
    {
        // UNC path to the workbook on a network share.
        string uncPath = @"\\ServerName\SharedFolder\MyWorkbook.xlsx";

        // Verify that the file exists before attempting to load.
        if (!File.Exists(uncPath))
        {
            Console.WriteLine($"File not found: {uncPath}");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(uncPath);

            // Example usage: print the name of the first worksheet.
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine("First worksheet name: " + firstSheet.Name);
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or processing.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
