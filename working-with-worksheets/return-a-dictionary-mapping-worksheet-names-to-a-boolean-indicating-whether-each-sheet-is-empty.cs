// Title: C# example: map worksheet names to boolean flags that show if each sheet is empty with Aspose.Cells
// Common Searches: c# Aspose.Cells find empty worksheets in an Excel workbook | how to generate a dictionary of sheet names and empty status using Aspose.Cells | determine if a worksheet has any data with MaxDataRow MaxDataColumn in C# | list Excel worksheets that contain no data using Aspose.Cells C# | check worksheet data presence programmatically with Aspose.Cells
// Tags: Aspose.Cells empty worksheet detection | C# map worksheet name to boolean | MaxDataRow MaxDataColumn usage Aspose.Cells | iterate Excel worksheets Aspose.Cells C# | Excel workbook empty sheet identification C# | dictionary of sheet emptiness Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file with Aspose.Cells, iterates through all worksheets, uses MaxDataRow and MaxDataColumn to determine if a sheet contains any data, stores each worksheet name with a true/false emptiness flag in a Dictionary<string, bool>, and outputs the results.
class Program
{
    static void Main(string[] args)
    {
        // Expect a single argument: the path to the Excel file.
        if (args.Length != 1)
        {
            Console.WriteLine("Usage: Program <excel-file-path>");
            return;
        }

        string filePath = args[0];

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found at '{filePath}'.");
            return;
        }

        try
        {
            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(filePath);

            // Dictionary to hold worksheet name -> isEmpty mapping.
            Dictionary<string, bool> worksheetEmptyMap = new Dictionary<string, bool>();

            // Iterate through each worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow and MaxDataColumn return -1 when there is no data.
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                bool isEmpty = (maxRow == -1) && (maxCol == -1);
                worksheetEmptyMap[sheet.Name] = isEmpty;
            }

            // Output the results.
            Console.WriteLine("Worksheet emptiness status:");
            foreach (var kvp in worksheetEmptyMap)
            {
                Console.WriteLine($"{kvp.Key}: {(kvp.Value ? "Empty" : "Not Empty")}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
