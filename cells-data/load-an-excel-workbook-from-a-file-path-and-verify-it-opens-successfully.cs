// Title: Load an Excel workbook from a specific file path and verify successful opening with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a spreadsheet file with Aspose.Cells, displays the total number of worksheets, and handles any load errors. | Create a reusable C# method that takes a file path, attempts to instantiate a Workbook via Aspose.Cells, and returns a boolean indicating whether the load succeeded. | Adapt the example to load a password‑protected .xlsx using Aspose.Cells and output either the worksheet count or the caught exception message.
// Common Searches: how to load an Excel workbook from a specific path using Aspose.Cells in C# | Aspose.Cells C# example for opening a workbook and checking if it opened successfully | C# catch exceptions when opening an .xlsx file with Aspose.Cells | retrieve number of worksheets after loading a workbook with Aspose.Cells .NET
// Tags: initialize Workbook with file path Aspose.Cells C# | check workbook load result Aspose.Cells | catch workbook loading errors Aspose.Cells | retrieve worksheet count Aspose.Cells | log workbook open status Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates loading an Excel workbook from a given file path using Aspose.Cells for .NET, prints a success message with the worksheet count, and catches any exceptions that occur during the load process.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = @"C:\Path\To\Your\File.xlsx";

        try
        {
            // Load the workbook from the specified file path
            Workbook workbook = new Workbook(filePath);

            // If no exception is thrown, the workbook opened successfully
            Console.WriteLine("Workbook loaded successfully.");
            Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading
            Console.WriteLine("Failed to load workbook: " + ex.Message);
        }
    }
}
