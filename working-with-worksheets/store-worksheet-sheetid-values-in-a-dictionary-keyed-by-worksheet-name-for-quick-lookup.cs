// Title: Create a C# dictionary that maps worksheet names to their internal SheetId indexes using Aspose.Cells
// AI Prompts: Write a C# function that loads an Excel workbook with Aspose.Cells and returns a Dictionary<string, int> where each key is a worksheet name and each value is the worksheet's SheetId. | Enhance the worksheet‑to‑SheetId mapper to also capture hidden worksheets and log any loading exceptions to a file. | Develop a console app that accepts a file path, uses Aspose.Cells to generate the name‑to‑SheetId map, and prints each worksheet and its SheetId to the console.
// Common Searches: C# Aspose.Cells get worksheet index by name | How to create a name-to-SheetId map for Excel worksheets in .NET | Retrieve internal SheetId of each sheet using Aspose.Cells | Dictionary of worksheet names and their indexes with Aspose.Cells C# example
// Tags: aspocells map worksheet name to sheetid | c# dictionary of excel sheet indexes | load workbook and enumerate worksheets aspocells | include hidden worksheets aspocells | exception handling for workbook load aspocells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// Provides a C# method that loads an Excel file with Aspose.Cells, iterates through all worksheets, and builds a Dictionary<string,int> mapping each worksheet's name to its internal SheetId (index). Includes file‑existence verification and basic exception logging.
class WorksheetIdMapper
{
    // Returns a dictionary where the key is the worksheet name
    // and the value is the worksheet's internal SheetId (index).
    public static Dictionary<string, int> GetWorksheetIdMap(string filePath)
    {
        var sheetIdMap = new Dictionary<string, int>();

        try
        {
            // Verify that the file exists before attempting to load.
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}");

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Store the sheet name as the key and its SheetId (index) as the value.
                sheetIdMap[sheet.Name] = sheet.Index;
            }
        }
        catch (Exception ex)
        {
            // Log the error and continue with an empty map.
            Console.Error.WriteLine($"Error loading workbook: {ex.Message}");
        }

        return sheetIdMap;
    }
}

class Program
{
    static void Main(string[] args)
    {
        // Use the first argument as the file path, or a default placeholder.
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        var worksheetMap = WorksheetIdMapper.GetWorksheetIdMap(filePath);

        // Output the mapping.
        foreach (var kvp in worksheetMap)
        {
            Console.WriteLine($"Worksheet: {kvp.Key}, SheetId: {kvp.Value}");
        }
    }
}
