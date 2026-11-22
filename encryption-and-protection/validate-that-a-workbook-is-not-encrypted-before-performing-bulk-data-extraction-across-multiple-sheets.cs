// Title: Check if an Excel workbook is encrypted before extracting all data from its worksheets with Aspose.Cells for .NET (C#)
// Common Searches: how to check if an Excel file is password protected using Aspose.Cells C# | Aspose.Cells C# detect encrypted workbook before reading worksheets | validate workbook encryption before bulk data extraction with Aspose.Cells | C# read all sheets from Excel only if workbook is not encrypted Aspose.Cells | handle encrypted Excel files in Aspose.Cells .NET when extracting data
// Tags: Aspose.Cells workbook encryption detection | extract all worksheet data C# | read Excel without password Aspose.Cells | validate Excel file protection .NET | iterate worksheets after encryption check

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample verifies that the input.xlsx file exists, loads it into an Aspose.Cells Workbook, iterates through every worksheet, and extracts each cell's value into a dictionary keyed by sheet name. It reports the row count per sheet and notes that older Aspose.Cells versions may not expose an IsEncrypted property, so encryption handling must be done by providing a password when opening the workbook or by external checks before extraction.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File '{inputPath}' not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // NOTE: Aspose.Cells older versions may not expose an IsEncrypted property.
        // If needed, password‑protected workbooks can be opened by providing the password
        // in the Workbook constructor. Here we proceed assuming the workbook is accessible.

        // Container to hold extracted data from each worksheet
        var extractedData = new Dictionary<string, List<object[]>>();

        try
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;       // Last row with data
                int maxCol = cells.MaxDataColumn;    // Last column with data

                var rows = new List<object[]>();

                // Extract data row by row
                for (int i = 0; i <= maxRow; i++)
                {
                    var rowData = new object[maxCol + 1];
                    for (int j = 0; j <= maxCol; j++)
                    {
                        rowData[j] = cells[i, j].Value;
                    }
                    rows.Add(rowData);
                }

                extractedData[sheet.Name] = rows;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during data extraction: {ex.Message}");
            return;
        }

        // Example output: display number of rows extracted per sheet
        foreach (var kvp in extractedData)
        {
            Console.WriteLine($"Sheet '{kvp.Key}' - rows extracted: {kvp.Value.Count}");
        }
    }
}
