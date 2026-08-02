// Title: C# Aspose.Cells – prepend "Reviewed:" to every cell value (bulk update)
// Description: This .NET example creates a workbook, adds sample data, then loops through the worksheet's Cells collection (as a FindAll substitute) to prepend the text "Reviewed:" to each non‑empty cell and saves the result as ReviewedOutput.xlsx.
// Keywords: Aspose.Cells | C# | .NET | prepend text to cells | bulk cell update | iterate over cells | Excel automation | FindAll alternative | GitHub sample | code snippet | Aspose.Cells API
// Common Searches: Add a prefix to all cells in an Aspose.Cells workbook C# | Bulk update cell values with Aspose.Cells .NET | Iterate over worksheet cells and modify content Aspose.Cells | How to prepend text to every Excel cell using Aspose.Cells | Aspose.Cells FindAll equivalent for updating cells
// Developer Intent: Add the prefix "Reviewed:" to each populated cell in a worksheet.
// Use Cases: Flag every entry as reviewed after data validation | Automatically tag cells with a status label before exporting reports | Apply a uniform prefix to all cells for branding or audit trails
// AI Prompts: Write C# code using Aspose.Cells to prepend 'Reviewed:' to all non‑empty cells and save the workbook. | Show how to simulate FindAll in Aspose.Cells by iterating over Cells and updating values. | Provide a GitHub‑ready snippet that bulk‑updates Excel cells with a custom prefix using Aspose.Cells .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsFindAllDemo
{
    // This .NET example creates a workbook, adds sample data, then loops through the worksheet's Cells collection (as a FindAll substitute) to prepend the text "Reviewed:" to each non‑empty cell and saves the result as ReviewedOutput.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // new Workbook("input.xlsx") to load

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Example data – populate some cells for demonstration
            cells["A1"].PutValue("Item 1");
            cells["B2"].PutValue("Item 2");
            cells["C3"].PutValue(123); // numeric value will be converted to string

            // Iterate over all cells (simulating FindAll) and prepend "Reviewed:"
            foreach (Cell cell in cells)
            {
                // Ensure the cell has a value
                if (cell.Value != null)
                {
                    // Get the current string representation of the cell's value
                    string currentValue = cell.StringValue;

                    // Prepend the prefix and write back to the cell
                    cell.PutValue("Reviewed:" + currentValue);
                }
            }

            // Save the workbook
            workbook.Save("ReviewedOutput.xlsx");
        }
    }
}
