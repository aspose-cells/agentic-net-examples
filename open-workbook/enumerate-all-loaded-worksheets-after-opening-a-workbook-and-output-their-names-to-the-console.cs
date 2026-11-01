// Title: List all worksheet names from an opened Excel file and write them to the console using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells read worksheet names from an existing workbook | How to print Excel sheet names to console with Aspose.Cells .NET | Iterate over worksheets in a loaded workbook using Aspose.Cells C# example | Handle password protected workbook while enumerating sheets with Aspose.Cells | Aspose.Cells get worksheet collection and display names in a console application
// Tags: enumerate worksheet names Aspose.Cells | iterate workbook worksheets .NET | print Excel sheet names console C# | load workbook and list sheets Aspose.Cells | password protected workbook exception Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the file "input.xlsx" exists, loads it into an Aspose.Cells Workbook, iterates through the Workbook.Worksheets collection, and writes each Worksheet.Name to the console. It includes specific handling for password‑protected workbooks via CellsException and a generic catch for other runtime errors.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be opened
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Output the name of each worksheet to the console
                Console.WriteLine(sheet.Name);
            }
        }
        catch (CellsException ex)
        {
            // Specific handling for password‑protected workbooks
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password to open it.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // General exception handling for any other runtime errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
