// Title: C# – Locate #DIV/0! Errors in a Named Range using Aspose.Cells FindOptions
// Description: Demonstrates how to create a workbook, define a named range, generate a #DIV/0! error, build a matching CellArea, configure FindOptions to search cell values, and retrieve the error cell with Cells.Find in Aspose.Cells for .NET.
// Keywords: Aspose.Cells FindOptions | C# error search | named range lookup | cell error detection | LookInType.Values | Find #DIV/0! Aspose
// Common Searches: Aspose.Cells find error cells in a specific range | C# locate #DIV/0! using FindOptions | search for error values inside a named range Aspose | how to limit Cells.Find to a CellArea
// Developer Intent: Identify cells that contain the #DIV/0! error within a predefined named range.
// Use Cases: Audit financial worksheets for division‑by‑zero issues before distribution. | Generate a log of error‑prone cells in a data‑import section for debugging. | Trigger conditional formatting only on error cells located in a targeted area.
// AI Prompts: Write C# code that uses Aspose.Cells FindOptions to retrieve all cells with any Excel error (e.g., #DIV/0!, #VALUE!) inside a given named range. | Explain the steps to set LookInType.Values and LookAtType.EntireContent for error‑string searches in a CellArea. | Provide a concise tutorial for extracting a named range, constructing the corresponding CellArea, and finding error cells with Cells.Find.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsFindErrorInNamedRange
{
    // Demonstrates how to create a workbook, define a named range, generate a #DIV/0! error, build a matching CellArea, configure FindOptions to search cell values, and retrieve the error cell with Cells.Find in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells – A2 will contain a division by zero formula (#DIV/0! error)
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].Formula = "=A1/0";   // This will produce #DIV/0!
                sheet.Cells["A3"].PutValue(5);

                // Define a named range that covers the cells we want to search
                // The range includes A1:A3 and is named "ErrorRange"
                sheet.Cells.CreateRange("A1", "A3").Name = "ErrorRange";

                // Retrieve the named range as a Range object (use alias to avoid conflict with System.Range)
                AsposeRange namedRange = workbook.Worksheets.GetRangeByName("ErrorRange");

                // Build a CellArea representing the same range (required by FindOptions)
                CellArea searchArea = new CellArea
                {
                    StartRow = namedRange.FirstRow,
                    StartColumn = namedRange.FirstColumn,
                    EndRow = namedRange.FirstRow + namedRange.RowCount - 1,
                    EndColumn = namedRange.FirstColumn + namedRange.ColumnCount - 1
                };

                // Configure FindOptions to search within the defined range for the exact error string
                FindOptions options = new FindOptions
                {
                    LookInType = LookInType.Values,          // Search in cell values (including errors)
                    LookAtType = LookAtType.EntireContent    // Match the whole cell content
                };
                options.SetRange(searchArea);                // Limit the search to the named range

                // Perform the search for the DIV/0 error string
                Cell errorCell = sheet.Cells.Find("#DIV/0!", null, options);

                // Output the result
                if (errorCell != null)
                {
                    Console.WriteLine($"Error cell found at: {errorCell.Name}");
                }
                else
                {
                    Console.WriteLine("No #DIV/0! error found in the named range.");
                }

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                workbook.Save("FindErrorInNamedRange.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
