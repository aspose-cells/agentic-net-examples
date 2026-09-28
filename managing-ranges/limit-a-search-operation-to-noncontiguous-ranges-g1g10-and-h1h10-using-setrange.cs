// Title: Limit Aspose.Cells Find search to non‑contiguous ranges G1:G10 and H1:H10 in C#
// AI Prompts: Use FindOptions.SetRange to locate the cell containing "G5" only within the ranges G1:G10 and H1:H10 in a C# worksheet. | Adapt the sample to search for any text across an array of CellArea objects that represent multiple disjoint ranges. | Write a reusable method that accepts a list of CellArea objects and a search string, then returns the first matching cell using Aspose.Cells Find.
// Common Searches: asp.net find value in specific non‑adjacent ranges using Aspose.Cells C# | how to restrict Aspose.Cells Find to multiple cell areas in C# | search only columns G and H with Aspose.Cells FindOptions SetRange | C# Aspose.Cells find across two separate ranges example
// Tags: Aspose.Cells FindOptions SetRange usage | non‑adjacent cell range search C# | search multiple CellArea objects Aspose.Cells | limit Excel find to specific columns Aspose | C# workbook find across disjoint ranges

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, fills columns G and H with sample data, defines two CellArea ranges (G1:G10 and H1:H10), and uses FindOptions.SetRange to search for the value "G5" first in the G column range and then in the H column range. It outputs the found cell address or a not‑found message and saves the file as SearchResult.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill sample data in columns G and H (indexes 6 and 7)
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 6].PutValue($"G{i + 1}");
                    sheet.Cells[i, 7].PutValue($"H{i + 1}");
                }

                // Define two separate areas: G1:G10 and H1:H10
                CellArea areaG = CellArea.CreateCellArea("G1", "G10");
                CellArea areaH = CellArea.CreateCellArea("H1", "H10");

                // Configure FindOptions
                FindOptions findOptions = new FindOptions();

                // Search in the first area
                findOptions.SetRange(areaG);
                // The Find method requires a start cell; passing null uses the default start cell
                Cell foundCell = sheet.Cells.Find("G5", null, findOptions);

                // If not found, search in the second area
                if (foundCell == null)
                {
                    findOptions.SetRange(areaH);
                    foundCell = sheet.Cells.Find("G5", null, findOptions);
                }

                // Output the result
                if (foundCell != null)
                {
                    Console.WriteLine($"Found at: {foundCell.Name}");
                }
                else
                {
                    Console.WriteLine("Value not found in the specified ranges.");
                }

                // Save the workbook (optional)
                workbook.Save("SearchResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
