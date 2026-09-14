// Title: Load an XLSX file from a local path with Aspose.Cells in C# and create a new workbook if the file is missing
// AI Prompts: Create C# code that uses Aspose.Cells to read an Excel workbook located at a specified path, verifies the file's presence, and if absent, instantiates a new Workbook and saves it to that location. | Write a try‑catch block in C# that loads an existing .xlsx using the Workbook(string) constructor and falls back to a default Workbook() with a save operation when the specified file cannot be found.
// Common Searches: c# aspocells load excel file using full path and create if missing | how to verify excel file existence before opening with Aspose.Cells in .NET | using Aspose.Cells Workbook constructor to open existing xlsx or generate new one | example handling missing Excel file when loading with Aspose.Cells C#
// Tags: Aspose.Cells open workbook from file path | C# verify Excel file existence before loading | fallback create new workbook Aspose.Cells | Aspose.Cells workbook initialization from file | handle missing XLSX file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks whether a given .xlsx file exists, loads it with the Workbook(string) constructor when present, otherwise creates a new Workbook, saves it to the same path, and wraps the process in a try‑catch for robust error handling.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = @"C:\path\to\your\file.xlsx";

        try
        {
            Workbook workbook;

            if (File.Exists(filePath))
            {
                // Load existing workbook
                workbook = new Workbook(filePath);
                Console.WriteLine("Workbook loaded successfully.");
            }
            else
            {
                // Create a new workbook as fallback and save it
                workbook = new Workbook();
                workbook.Save(filePath);
                Console.WriteLine($"File not found. Created new workbook and saved to '{filePath}'.");
            }

            // Further operations on 'workbook' can be added here
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
