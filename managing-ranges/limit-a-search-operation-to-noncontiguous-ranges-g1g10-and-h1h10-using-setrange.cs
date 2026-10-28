// Title: How to search for a value in non‑contiguous ranges G1:G10 and H1:H10 using Aspose.Cells FindOptions SetRange in C#
// Common Searches: Aspose.Cells FindOptions SetRange example for searching multiple separate ranges in C# | C# code to locate a cell value only within G1:G10 and H1:H10 with Aspose.Cells | How to limit Aspose.Cells Find method to specific cell areas like column G and column H | Search Excel workbook for a string in two non‑adjacent columns using Aspose.Cells C# | Find value in non‑contiguous ranges using Aspose.Cells Find with SetRange
// Tags: Aspose.Cells FindOptions SetRange C# | search non‑contiguous cell ranges Aspose.Cells | limit Find to specific columns G H Aspose.Cells | C# Excel find value across multiple ranges | Aspose.Cells find across separate cell areas

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, fills columns G and H (rows 1‑10) with data, places the string "Target" in G5, and demonstrates how to use FindOptions with SetRange to search for "Target" first in the G1:G10 range and then in the H1:H10 range. It outputs the address of the found cell and saves the workbook as SearchResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in G1:G10 and H1:H10
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 6].PutValue($"G{i + 1}"); // Column G (index 6)
                sheet.Cells[i, 7].PutValue($"H{i + 1}"); // Column H (index 7)
            }

            // Place the target value somewhere within the defined ranges
            sheet.Cells["G5"].PutValue("Target");

            // Prepare FindOptions for the first range G1:G10
            FindOptions options = new FindOptions();
            CellArea rangeG = CellArea.CreateCellArea("G1", "G10");
            options.SetRange(rangeG);

            // Search in the first range
            Cell foundCell = sheet.Cells.Find("Target", null, options);

            // If not found, search in the second range H1:H10
            if (foundCell == null)
            {
                CellArea rangeH = CellArea.CreateCellArea("H1", "H10");
                options.SetRange(rangeH);
                foundCell = sheet.Cells.Find("Target", null, options);
            }

            // Output result
            if (foundCell != null)
            {
                Console.WriteLine($"Found at {foundCell.Name}");
            }
            else
            {
                Console.WriteLine("Not found in the specified ranges.");
            }

            // Save the workbook (ensure the directory exists)
            string outputPath = "SearchResult.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
