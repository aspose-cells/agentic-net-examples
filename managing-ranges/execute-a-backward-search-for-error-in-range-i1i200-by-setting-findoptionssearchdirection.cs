using System;
using Aspose.Cells;

namespace AsposeCellsFindBackwardDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Populate sample data in column I (index 8) rows 1-200
            for (int row = 0; row < 200; row++)
            {
                // Example: put the word "Error" in every 10th row
                if (row % 10 == 0)
                    worksheet.Cells[row, 8].PutValue("Error");
                else
                    worksheet.Cells[row, 8].PutValue($"Info_{row}");
            }

            // Create FindOptions and configure it for a backward search
            FindOptions options = new FindOptions
            {
                LookInType = LookInType.Values,          // Search in cell values
                LookAtType = LookAtType.Contains,        // Match if the cell contains the key
                SearchBackward = true                    // Enable backward search
            };

            // Define the search range I1:I200 (rows 0-199, column 8)
            CellArea searchRange = new CellArea
            {
                StartRow = 0,
                StartColumn = 8,
                EndRow = 199,
                EndColumn = 8
            };
            options.SetRange(searchRange);

            // Perform the find operation for the text "Error"
            Cell foundCell = worksheet.Cells.Find("Error", null, options);

            // Output the result
            if (foundCell != null)
                Console.WriteLine($"Found \"Error\" at cell {foundCell.Name}");
            else
                Console.WriteLine("The text \"Error\" was not found in the specified range.");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("FindBackwardResult.xlsx");
        }
    }
}