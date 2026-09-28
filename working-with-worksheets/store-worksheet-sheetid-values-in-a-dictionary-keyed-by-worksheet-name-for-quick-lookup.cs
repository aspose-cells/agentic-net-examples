// Title: Build a case‑insensitive dictionary that maps worksheet names to their SheetId (zero‑based index) using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells and creates a Dictionary<string,int> where each key is the worksheet name (case‑insensitive) and each value is the worksheet's Index. | Show how to retrieve a worksheet's Index from the dictionary by name and handle the situation when the name does not exist. | Add comprehensive error handling for missing input file, workbook load failures, and unexpected exceptions while populating the name‑to‑index map.
// Common Searches: aspnet c# create dictionary of worksheet names to sheet indexes using Aspose.Cells | case insensitive lookup of Excel sheet index by name Aspose.Cells | how to map worksheet Name to Index in Aspose.Cells workbook | retrieve sheet Id from worksheet name Aspose.Cells .NET example | error handling when loading workbook and building sheet name dictionary Aspose.Cells
// Tags: worksheet name to index dictionary Aspose.Cells | case‑insensitive sheet lookup C# | Aspose.Cells workbook sheet mapping | Excel sheet index retrieval Aspose.Cells | robust workbook load error handling Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// // Loads an Excel file with Aspose.Cells, iterates through all worksheets, and builds a case‑insensitive Dictionary<string,int> that maps each worksheet's Name to its zero‑based Index. Demonstrates safe retrieval of a sheet's Index and includes error handling for missing files and load failures.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Map worksheet name to its index (zero‑based)
        var sheetIdByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        try
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Worksheet.Index uniquely identifies the sheet within the workbook
                sheetIdByName[sheet.Name] = sheet.Index;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while processing worksheets: {ex.Message}");
            return;
        }

        // Example: retrieve index for a given worksheet name
        if (sheetIdByName.TryGetValue("Sheet1", out int sheetId))
        {
            Console.WriteLine($"Sheet1 has Index = {sheetId}");
        }
        else
        {
            Console.WriteLine("Sheet1 not found.");
        }
    }
}
