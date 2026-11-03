// Title: Sort an Excel column of month names with a custom month order using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells sort column by custom month sequence | apply predefined month ordering to Excel data with Aspose.Cells .NET | how to implement custom list sorting for month names in Aspose.Cells | using Array.IndexOf for custom month sort in an Aspose.Cells workbook | sort rows based on a month array in C# Excel library
// Tags: custom month ordering Aspose.Cells | sort column C by predefined month list | Aspose.Cells custom list sort .NET | C# Excel month array sorting | write sorted months to worksheet Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The example creates a workbook, fills column C with month names, defines a custom month order array, extracts the month values, sorts the rows according to that order using Array.IndexOf, writes the sorted months back to column C, and saves the file as SortedMonths.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column C (index 2)
            sheet.Cells["C1"].PutValue("Month");
            sheet.Cells["C2"].PutValue("March");
            sheet.Cells["C3"].PutValue("January");
            sheet.Cells["C4"].PutValue("February");
            sheet.Cells["C5"].PutValue("December");

            // Define custom month order
            string[] monthOrder = new[]
            {
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            };

            // Collect month values from rows 2 to 5 (zero‑based indices 1‑4)
            int startRow = 1; // C2
            int endRow = 4;   // C5
            List<(int RowIndex, string Month)> rows = new List<(int, string)>();

            for (int i = startRow; i <= endRow; i++)
            {
                string month = sheet.Cells[i, 2].StringValue; // column C (index 2)
                rows.Add((i, month));
            }

            // Sort rows according to the custom month order
            List<(int RowIndex, string Month)> sortedRows = rows
                .OrderBy(r => Array.IndexOf(monthOrder, r.Month))
                .ToList();

            // Write sorted month values back to the worksheet
            for (int i = 0; i < sortedRows.Count; i++)
            {
                sheet.Cells[startRow + i, 2].PutValue(sortedRows[i].Month);
            }

            // Define output file path
            string outputPath = "SortedMonths.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
