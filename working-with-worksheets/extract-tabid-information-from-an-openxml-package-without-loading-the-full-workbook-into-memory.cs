// Title: Retrieve each worksheet's TabId from an .xlsx file using Aspose.Cells in C#
// AI Prompts: Generate a C# method that accepts a .xlsx file path and returns a Dictionary<string,long> mapping each worksheet name to its TabId, using Aspose.Cells. | Show how to iterate over Workbook.Worksheets in Aspose.Cells and capture the TabId property for every sheet, including a file‑existence check. | Provide sample code that safely handles exceptions while reading worksheet TabIds from an Excel file with Aspose.Cells. | Demonstrate how to print the collected sheet‑name‑to‑TabId pairs to the console in C#.
// Common Searches: Aspose.Cells C# get worksheet TabId dictionary | How to read Excel sheet TabId with Aspose.Cells in .NET | C# code to list all sheet names and their TabId from .xlsx | Retrieve TabId values from worksheets without opening Excel | Aspose.Cells example for extracting TabId property from worksheets
// Tags: Aspose.Cells read worksheet TabId | C# extract sheet TabId from .xlsx | Aspose.Cells enumerate workbook worksheets | C# dictionary mapping sheet name to TabId | error‑handling Aspose.Cells file read | Aspose.Cells workbook TabId property

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx workbook with Aspose.Cells, loops through each worksheet, and builds a Dictionary that maps worksheet names to their TabId values, while handling missing files and runtime exceptions.
public static class ExcelTabIdExtractor
{
    /// <param name="filePath">Full path to the .xlsx file.</param>
    /// <returns>Dictionary where the key is the sheet name and the value is its TabId.</returns>
    public static Dictionary<string, long> GetSheetTabIds(string filePath)
    {
        var sheetTabIds = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);

        // Prevent FileNotFoundException.
        if (!File.Exists(filePath))
            return sheetTabIds;

        try
        {
            // Load the workbook using Aspose.Cells.
            var workbook = new Workbook(filePath);

            // Iterate through worksheets and collect their TabId.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // TabId is an integer; cast to long for consistency with the return type.
                sheetTabIds[sheet.Name] = (long)sheet.TabId;
            }
        }
        catch (Exception)
        {
            // Runtime safety: swallow exceptions or handle logging as needed.
        }

        return sheetTabIds;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Example usage: first argument is the path to the .xlsx file.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to an .xlsx file as a command‑line argument.");
                return;
            }

            string filePath = args[0];
            var tabIds = ExcelTabIdExtractor.GetSheetTabIds(filePath);

            if (tabIds.Count == 0)
            {
                Console.WriteLine("No sheet TabIds were found or the file could not be processed.");
                return;
            }

            Console.WriteLine("Sheet TabIds:");
            foreach (var kvp in tabIds)
            {
                Console.WriteLine($"Sheet: {kvp.Key}, TabId: {kvp.Value}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
