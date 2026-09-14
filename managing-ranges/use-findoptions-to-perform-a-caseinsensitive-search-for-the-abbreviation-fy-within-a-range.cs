// Title: Use Aspose.Cells FindOptions for a case‑insensitive search of the abbreviation “FY” within cells A1:A5 in C#
// AI Prompts: Locate the first cell containing the text "FY" without regard to case in rows 1‑5 of column A using ws.Cells.Find and output its address. | Set up search options with CaseSensitive = false and LookInType.Values to scan a defined range, then confirm the result falls inside A1:A5. | After the match is found, save the workbook as "FindResult.xlsx" while preserving the configured search behavior.
// Common Searches: aspnet find abbreviation FY in Excel range A1:A5 ignoring case with Aspose.Cells | c# Aspose.Cells search column A for text FY and ignore case | limit ws.Cells.Find to cells A1:A5 and perform ignore case lookup | use LookInType.Values with FindOptions to locate text in an Excel file via C# | verify that Cells.Find result is within a specific range in Aspose.Cells
// Tags: case‑insensitive FindOptions search Aspose.Cells | Cells.Find search cell values C# | find abbreviation FY in Excel range Aspose | search values only using LookInType Aspose.Cells | validate found cell within specific range C# | Excel range A1:A5 search Aspose.Cells

using Aspose.Cells;
using System;

// // Demonstrates using Aspose.Cells FindOptions to perform a case‑insensitive search for "FY" within cells A1:A5, validates the located cell, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("FY2021");
            ws.Cells["A2"].PutValue("fy2022");
            ws.Cells["A3"].PutValue("Q1");
            ws.Cells["A4"].PutValue("fy");
            ws.Cells["A5"].PutValue("FY");

            // Configure FindOptions for a case‑insensitive search
            FindOptions options = new FindOptions
            {
                CaseSensitive = false,               // ignore case
                LookInType = LookInType.Values       // search cell values
            };

            // Start the search from the first cell (A1)
            Cell startCell = ws.Cells["A1"];

            // Perform the search for the abbreviation "FY"
            Cell foundCell = ws.Cells.Find("FY", startCell, options);

            // Verify that the found cell lies within the desired range A1:A5
            if (foundCell != null && foundCell.Row >= 0 && foundCell.Row <= 4 && foundCell.Column == 0)
            {
                Console.WriteLine($"Found at {foundCell.Name}: {foundCell.StringValue}");
            }
            else
            {
                Console.WriteLine("Abbreviation 'FY' not found in the range.");
            }

            // Save the workbook (optional)
            wb.Save("FindResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
