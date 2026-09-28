// Title: Find the text "Total" in cells A1:C30 using Aspose.Cells FindOptions in C#
// AI Prompts: Use FindOptions to locate the cell that contains the word "Total" within the range A1:C30 and return its address. | Modify the code to search for a different keyword (e.g., "Subtotal") while keeping the search limited to the same A1:C30 range. | Change LookInType to LookInType.Formulas to search for "Total" inside formulas in the defined range.
// Common Searches: Aspose.Cells C# find specific text inside a defined range with FindOptions | how to limit Aspose.Cells Find search to cells A1:C30 | validate that a cell found by Find is inside a given range using Aspose.Cells | search for a string in Excel worksheet range using Aspose.Cells FindOptions C# | Aspose.Cells LookInType values vs formulas example
// Tags: Aspose.Cells FindOptions search within range | C# locate text in Excel cells using Aspose.Cells | validate found cell inside specified range Aspose.Cells | Aspose.Cells LookInType values search | Excel cell address retrieval Aspose.Cells C#

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example creates a workbook, defines the range A1:C30, configures FindOptions to look in cell values, searches for the text "Total" starting from the top‑left cell of the range, verifies that the found cell lies within the defined range, and outputs the cell address or a not‑found message.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the search range A1:C30
                AsposeRange searchRange = worksheet.Cells.CreateRange("A1:C30");

                // Starting cell for the search (top‑left cell of the range)
                Cell startCell = worksheet.Cells[searchRange.FirstRow, searchRange.FirstColumn];

                // Set up FindOptions to search within cell values
                FindOptions findOptions = new FindOptions
                {
                    LookInType = LookInType.Values
                };

                // Perform the search for the text "Total"
                Cell foundCell = worksheet.Cells.Find("Total", startCell, findOptions);

                // Verify that the found cell lies within the defined range
                bool withinRange = false;
                if (foundCell != null)
                {
                    int lastRow = searchRange.FirstRow + searchRange.RowCount - 1;
                    int lastColumn = searchRange.FirstColumn + searchRange.ColumnCount - 1;
                    withinRange = foundCell.Row >= searchRange.FirstRow && foundCell.Row <= lastRow &&
                                  foundCell.Column >= searchRange.FirstColumn && foundCell.Column <= lastColumn;
                }

                // Output the result
                if (foundCell != null && withinRange)
                {
                    Console.WriteLine($"Found \"Total\" at cell: {foundCell.Name}");
                }
                else
                {
                    Console.WriteLine("The text \"Total\" was not found in the specified range.");
                }

                // (Optional) Save the workbook if needed
                // workbook.Save("Output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
