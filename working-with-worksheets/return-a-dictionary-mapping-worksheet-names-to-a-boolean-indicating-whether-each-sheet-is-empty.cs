// Title: Get a dictionary of worksheet names with empty‑sheet flags using Aspose.Cells in C#
// AI Prompts: Write a C# method that loads an Excel file with Aspose.Cells and returns a Dictionary<string,bool> where each key is a worksheet name and the value is true if the sheet has no data. | Add robust error handling so the method logs any loading exceptions and returns an empty dictionary when the workbook cannot be opened. | Extend the method to also output the count of populated rows for each non‑empty worksheet alongside the emptiness flag.
// Common Searches: how to check if an Excel worksheet is empty using Aspose.Cells C# | C# Aspose.Cells get worksheet names and empty status from a workbook | retrieve list of empty sheets from Excel file with Aspose.Cells .NET | Aspose.Cells MaxDataRow MaxDataColumn empty sheet detection C#
// Tags: Aspose.Cells empty worksheet detection | Aspose.Cells MaxDataRow usage | Aspose.Cells iterate workbook worksheets | C# load Excel workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates each worksheet, uses Cells.MaxDataRow and Cells.MaxDataColumn to determine if the sheet contains any data, and returns a Dictionary mapping worksheet names to a boolean that indicates whether the sheet is empty.
public class WorksheetHelper
{
    /// <param name="filePath">Full path to the Excel file to be examined.</param>
    /// <returns>Dictionary with worksheet names as keys and emptiness flag as values.</returns>
    public static Dictionary<string, bool> GetWorksheetEmptyStatus(string filePath)
    {
        var sheetEmptyMap = new Dictionary<string, bool>();

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine if the sheet contains any data.
                // MaxDataRow and MaxDataColumn return -1 when there is no data.
                bool isEmpty = sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0;

                // Add the result to the dictionary using the sheet's name.
                sheetEmptyMap[sheet.Name] = isEmpty;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            // Return an empty dictionary on failure.
        }

        return sheetEmptyMap;
    }

    // Entry point for the console application.
    public static void Main(string[] args)
    {
        // Determine the Excel file path: first argument or a default placeholder.
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        // Verify that the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Retrieve emptiness status for each worksheet.
            Dictionary<string, bool> statusMap = GetWorksheetEmptyStatus(filePath);

            // Output the results.
            foreach (KeyValuePair<string, bool> kvp in statusMap)
            {
                Console.WriteLine($"{kvp.Key}: {(kvp.Value ? "Empty" : "Not Empty")}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
