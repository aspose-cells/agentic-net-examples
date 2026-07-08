using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Two‑dimensional array of booleans to import
        bool[,] boolData = new bool[,]
        {
            { true,  false, true  },
            { false, true,  false }
        };

        // Convert the bool[,] to object[,] because ImportTwoDimensionArray expects an object array
        int rows = boolData.GetLength(0);
        int cols = boolData.GetLength(1);
        object[,] objData = new object[rows, cols];
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                objData[r, c] = boolData[r, c];

        // Import the two‑dimensional array starting at cell A1 (lifecycle rule: import)
        cells.ImportTwoDimensionArray(objData, 0, 0);

        // Apply checkbox style to each imported cell and ensure the boolean value is set
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                Cell cell = cells[r, c];
                cell.IsCheckBoxStyle = true;          // display as a check box
                cell.PutValue(boolData[r, c]);        // true shows a check mark, false shows empty
            }
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save("BooleanCheckBox.xlsx");
    }
}