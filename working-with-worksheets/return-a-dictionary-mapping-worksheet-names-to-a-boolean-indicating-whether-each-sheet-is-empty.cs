// Title: C# – Get a Dictionary of Worksheet Names with Empty‑Sheet Flags using Aspose.Cells
// Description: Sample C# code that loads an Excel workbook with Aspose.Cells for .NET, iterates every worksheet, and marks it as empty when MaxDataRow or MaxDataColumn is –1. The method returns a Dictionary<string, bool> where the key is the sheet name and the value indicates emptiness, with basic error handling and console output.
// Keywords: Aspose.Cells empty worksheet detection | C# Aspose.Cells GetEmptySheetMap | MaxDataRow MaxDataColumn empty sheet | dictionary of sheet names Aspose | .NET Excel empty sheet check | Aspose.Cells sample code GitHub | Excel worksheet validation C# | detect blank worksheets Aspose.Cells
// Common Searches: how to check if an Excel sheet is empty with Aspose.Cells | C# get list of empty worksheets in a workbook | Aspose.Cells dictionary of sheet name to empty flag | detect blank worksheets using MaxDataRow | Aspose.Cells sample for empty sheet detection
// Developer Intent: Identify which worksheets in an Excel file contain no data and return a name‑to‑boolean map.
// Use Cases: Validate template files and skip completely blank sheets before data import. | Generate a cleanup report that lists empty worksheets for end‑users. | Improve processing speed by ignoring worksheets with no rows or columns of data.
// AI Prompts: Write C# code with Aspose.Cells that returns a Dictionary<string,bool> indicating empty worksheets based on MaxDataRow and MaxDataColumn. | Explain how MaxDataRow and MaxDataColumn can be used to determine worksheet emptiness in Aspose.Cells for .NET. | Add robust logging and custom exceptions to the GetEmptySheetMap method for production use.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Sample C# code that loads an Excel workbook with Aspose.Cells for .NET, iterates every worksheet, and marks it as empty when MaxDataRow or MaxDataColumn is –1. The method returns a Dictionary<string, bool> where the key is the sheet name and the value indicates emptiness, with basic error handling and console output.
public class WorksheetEmptyChecker
{
    /// <param name="filePath">Path to the Excel file to be examined.</param>
    /// <returns>Dictionary where key = worksheet name, value = true if the sheet is empty.</returns>
    public static Dictionary<string, bool> GetEmptySheetMap(string filePath)
    {
        var sheetEmptyMap = new Dictionary<string, bool>();

        try
        {
            // Load the workbook
            var workbook = new Workbook(filePath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // A sheet is considered empty when it has no used rows or columns.
                // MaxDataRow and MaxDataColumn return -1 when there is no data.
                bool isEmpty = sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0;
                sheetEmptyMap[sheet.Name] = isEmpty;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }

        return sheetEmptyMap;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Determine the Excel file path (first argument or default)
        string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Get the empty sheet map
            Dictionary<string, bool> result = WorksheetEmptyChecker.GetEmptySheetMap(filePath);

            // Output the results
            foreach (var kvp in result)
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
