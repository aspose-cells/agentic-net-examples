// Title: How to set the font color of a pivot table row field header to black using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set pivot table row field header font color to black | Change font color of a pivot table header cell with Aspose.Cells .NET | Programmatically modify pivot table header style in C# | Set Excel pivot table header text color to black using Aspose.Cells API | Retrieve pivot field header cell and apply black font color in C#
// Tags: Aspose.Cells pivot header style API | C# modify pivot table header cell | Excel pivot row field formatting with Aspose | set pivot table header text color .NET | pivot table header cell style example

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, accesses the first worksheet and its first pivot table, locates the header cell of the first row field, changes the font color of that cell to black via the cell's style, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify the required one)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                    return;
                }

                // Assume the pivot table you want to modify is the first one on the sheet
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Ensure the pivot table has at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("Error: The pivot table does not contain any row fields.");
                    return;
                }

                // Retrieve the first row field (header) of the pivot table
                PivotField rowField = pivotTable.RowFields[0];

                // The header cell for a row field is located at the start of the RowRange
                // plus the field's position within the row area.
                int headerRow = pivotTable.RowRange.StartRow;
                int headerColumn = pivotTable.RowRange.StartColumn + rowField.Position;

                // Access the header cell
                Cell headerCell = worksheet.Cells[headerRow, headerColumn];

                // Change the font color of the header cell to black
                Style headerStyle = headerCell.GetStyle();
                headerStyle.Font.Color = Color.Black;
                headerCell.SetStyle(headerStyle);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
