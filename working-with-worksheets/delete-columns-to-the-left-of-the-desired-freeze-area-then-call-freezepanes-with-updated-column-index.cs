// Title: Delete Columns Left of Freeze Pane and Apply FreezePanes with Aspose.Cells (C#)
// Description: Creates a workbook, fills columns A‑E, removes all columns before a specified index, then freezes the first row and the new first column using FreezePanes. Optionally sets the first visible column of the right pane and saves the file.
// Keywords: Aspose.Cells C# delete columns | FreezePanes after column removal | Aspose.Cells set FirstVisibleColumnOfRightPane | C# Excel freeze rows and columns | Aspose.Cells worksheet manipulation
// Common Searches: Aspose.Cells delete columns before freeze pane | How to use FreezePanes after removing columns in .NET | Set first visible column of right pane Aspose.Cells | C# freeze top row and left column after column deletion | Aspose.Cells example FreezePanes with updated index
// Developer Intent: Remove unwanted columns and then correctly apply FreezePanes at the new column position using Aspose.Cells for .NET.
// Use Cases: Eliminate preceding columns so a target column becomes the leftmost visible column, then freeze it with the header row. | Maintain header visibility while scrolling horizontally and vertically after column cleanup. | Control which column appears first in the scrollable right pane by setting FirstVisibleColumnOfRightPane. | Generate an Excel file with a customized freeze layout for reporting or dashboard purposes.
// AI Prompts: Write C# code that deletes columns left of a given index and then calls FreezePanes with the adjusted column index using Aspose.Cells. | Explain how to configure FirstVisibleColumnOfRightPane after applying FreezePanes in an Aspose.Cells workbook. | Provide best‑practice error handling for column deletion and FreezePanes operations in Aspose.Cells C# examples.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills columns A‑E, removes all columns before a specified index, then freezes the first row and the new first column using FreezePanes. Optionally sets the first visible column of the right pane and saves the file.
    public class FreezePanesAfterDeleteColumns
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data in columns A‑E
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[0, col].PutValue($"Header {(char)('A' + col)}");
                    sheet.Cells[1, col].PutValue(col + 1);
                }

                // Desired freeze column (0‑based). Example: freeze at column D (index 3)
                int desiredFreezeColumn = 3;

                // Delete all columns to the left of the desired freeze column
                if (desiredFreezeColumn > 0)
                {
                    // DeleteColumns(startIndex, totalColumns, updateReference)
                    sheet.Cells.DeleteColumns(0, desiredFreezeColumn, true);
                }

                // After deletion the column we want to freeze becomes index 0
                int freezeRow = 1;      // Freeze first row (row index 1 = second row)
                int freezeColumn = 0;   // Freeze first column after deletion

                // Apply freeze panes
                // FreezePanes(row, column, freezedRows, freezedColumns)
                sheet.FreezePanes(freezeRow, freezeColumn, freezeRow, freezeColumn);

                // Optional: adjust the first visible column of the right pane
                PaneCollection panes = sheet.GetPanes();
                panes.FirstVisibleColumnOfRightPane = 1;

                // Save the workbook
                string outputPath = "FreezeAfterDeleteColumns.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            FreezePanesAfterDeleteColumns.Run();
        }
    }
}
