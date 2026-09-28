// Title: Identify Excel cells with numeric values above 1000 and export their addresses using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, scans every cell in a worksheet, and returns a list of cell names where the numeric value exceeds 1000. | Enhance the solution to accept a configurable numeric limit and handle decimal numbers, then write the matching cell addresses to a newly added worksheet. | Create a reusable method that takes input and output file paths plus a numeric threshold, extracts addresses of cells greater than that threshold, and saves them to a separate sheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# find cell addresses where value is greater than 1000 | How to list Excel cells with numbers over a certain threshold using Aspose.Cells | Export addresses of large numeric values from a worksheet with Aspose.Cells .NET | C# Aspose.Cells filter cells by numeric limit and write results to new sheet
// Tags: find cells with numeric value exceeding threshold Aspose.Cells | collect cell addresses from Excel using Aspose.Cells C# | write address list to new worksheet Aspose.Cells | iterate all worksheet cells Aspose.Cells .NET | filter large numbers in Excel with Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook, iterates through every cell in the first worksheet, records the addresses of numeric values (int, long, double) greater than 1000, writes those addresses to a new worksheet named "LargeValues", and saves the updated file.
class FindLargeValues
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Work with the first worksheet (you can loop through all worksheets if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // List to store addresses of cells whose numeric value is greater than 1000
        List<string> largeValueAddresses = new List<string>();

        // Iterate through all cells in the worksheet
        foreach (Cell cell in worksheet.Cells)
        {
            // Skip cells that have no value
            if (cell.Value == null) continue;

            // Check for double values
            if (cell.Value is double d && d > 1000)
            {
                largeValueAddresses.Add(cell.Name);
                continue;
            }

            // Check for integer values (Aspose may store integers as int or long)
            if (cell.Value is int i && i > 1000)
            {
                largeValueAddresses.Add(cell.Name);
                continue;
            }

            if (cell.Value is long l && l > 1000)
            {
                largeValueAddresses.Add(cell.Name);
                continue;
            }
        }

        // Write the collected addresses to a new worksheet for reference
        Worksheet resultSheet = workbook.Worksheets.Add("LargeValues");
        for (int i = 0; i < largeValueAddresses.Count; i++)
        {
            resultSheet.Cells[i, 0].PutValue(largeValueAddresses[i]);
        }

        // Save the workbook with the results (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
