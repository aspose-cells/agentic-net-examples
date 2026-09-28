// Title: Import a 2‑D boolean array into an Aspose.Cells worksheet and render true values as checkboxes
// AI Prompts: Write C# code that imports a two‑dimensional boolean array into a worksheet and enables the built‑in checkbox visual for cells containing true using Aspose.Cells. | Generate a program that reads a bool matrix, places it into Excel via Aspose.Cells, and marks each true entry with a checked box.
// Common Searches: Aspose.Cells how to show check boxes for true values from a boolean array | C# import 2D bool array into Excel and enable checkbox style with Aspose.Cells | set IsCheckBoxStyle on imported cells Aspose.Cells example | display check marks for true booleans in Excel using Aspose.Cells
// Tags: import boolean 2d array Aspose.Cells | set cell IsCheckBoxStyle C# | display check mark for true values Excel | boolean matrix to checkbox style | Aspose.Cells checkbox rendering

using System;
using Aspose.Cells;

namespace AsposeCellsBooleanCheckboxDemo
{
    // // Creates a workbook, imports a 2‑D object array of booleans starting at A1, iterates over each cell to set IsCheckBoxStyle = true so true values appear as check marks, and saves the file as BooleanCheckBox.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define a two‑dimensional array of booleans (as objects)
            object[,] boolData = new object[,]
            {
                { true, false, true },
                { false, true, false },
                { true, true, false }
            };

            // Import the 2‑D array starting at cell A1 (row 0, column 0)
            cells.ImportTwoDimensionArray(boolData, 0, 0);

            // Apply checkbox style to each imported cell
            int rows = boolData.GetLength(0);
            int cols = boolData.GetLength(1);
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    Cell cell = cells[r, c];
                    cell.IsCheckBoxStyle = true;   // display as a check box
                    // The boolean value is already in the cell; true shows a check mark
                }
            }

            // Save the workbook
            workbook.Save("BooleanCheckBox.xlsx");
        }
    }
}
