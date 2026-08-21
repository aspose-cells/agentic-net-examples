// Title: C# Aspose.Cells: Backward search for "Error" in range I1:I200 with FindOptions
// Description: This example creates a workbook, fills column I with sample data, defines the cell area I1:I200, and configures FindOptions (LookInType.Values, LookAtType.Contains, SearchBackward=true) to locate the last occurrence of the word "Error". The code demonstrates limiting the search to a specific range and retrieving the cell address.
// Keywords: Aspose.Cells | C# | FindOptions | SearchBackward | cell range I1:I200 | find text Error | worksheet.Cells.Find | backward search | .NET spreadsheet API
// Common Searches: Aspose.Cells backward search in a column | FindOptions limit search to specific range .NET | Locate last occurrence of a string in Excel using Aspose.Cells | Search cells containing 'Error' in column I with C#
// Developer Intent: Find the most recent cell that contains the word "Error" within column I (rows 1‑200) by searching the range in reverse order.
// Use Cases: Retrieve the latest error entry from a log stored in column I. | Confirm the presence of an error marker before processing subsequent rows. | Generate an audit note with the address of the last "Error" cell.
// AI Prompts: Show how to change the code to perform a forward search instead of a backward one. | Explain how to collect all cells that contain "Error" in the I1:I200 range. | Provide code to apply a red background to the cell found by the backward search.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, fills column I with sample data, defines the cell area I1:I200, and configures FindOptions (LookInType.Values, LookAtType.Contains, SearchBackward=true) to locate the last occurrence of the word "Error". The code demonstrates limiting the search to a specific range and retrieving the cell address.
    public class BackwardSearchErrorDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data in column I (index 8) for demonstration
                for (int row = 0; row < 200; row++)
                {
                    // Insert the word "Error" in some cells to be found
                    if (row % 25 == 0) // every 25th row contains "Error"
                    {
                        worksheet.Cells[row, 8].PutValue("Error");
                    }
                    else
                    {
                        worksheet.Cells[row, 8].PutValue($"Data_{row}");
                    }
                }

                // Define the search range I1:I200
                CellArea searchRange = new CellArea
                {
                    StartRow = 0,          // I1 -> row 0
                    StartColumn = 8,       // column I -> index 8
                    EndRow = 199,          // I200 -> row 199
                    EndColumn = 8
                };

                // Configure find options for a backward search
                FindOptions options = new FindOptions
                {
                    LookInType = LookInType.Values,   // search in cell values
                    LookAtType = LookAtType.Contains, // match if the cell contains the text
                    SearchBackward = true             // enable backward search
                };
                options.SetRange(searchRange);        // limit the search to I1:I200

                // Perform the search for the text "Error"
                Cell foundCell = worksheet.Cells.Find("Error", null, options);

                // Output the result
                if (foundCell != null)
                {
                    Console.WriteLine($"Found \"Error\" at cell {foundCell.Name} (Row {foundCell.Row + 1}, Column {foundCell.Column + 1})");
                }
                else
                {
                    Console.WriteLine("The text \"Error\" was not found in the specified range.");
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                workbook.Save("BackwardSearchErrorDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                BackwardSearchErrorDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
