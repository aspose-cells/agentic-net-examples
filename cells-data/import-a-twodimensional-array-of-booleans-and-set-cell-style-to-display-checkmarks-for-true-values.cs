// Title: Import a 2D Boolean array into Aspose.Cells and display true values as checkboxes (C#)
// Description: Demonstrates how to convert a bool[,] to an object[,] for ImportTwoDimensionArray, load the data into a worksheet starting at A1, apply the IsCheckBoxStyle property to each cell so that true entries appear as checkmarks, and save the workbook as BooleanCheckBox.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | ImportTwoDimensionArray | bool[,] to object[,] conversion | IsCheckBoxStyle | checkbox style Excel | boolean array import | display checkmarks | Excel workbook generation
// Common Searches: Aspose.Cells import boolean array C# | show true values as checkboxes in Excel | IsCheckBoxStyle property example | convert bool[,] for ImportTwoDimensionArray | checkbox style for imported cells
// Developer Intent: Load a two‑dimensional boolean array into a worksheet and format true cells as visual checkboxes.
// Use Cases: Render survey results where each selected answer is a checked box. | Create a feature‑flag matrix with checkmarks indicating enabled features. | Build a status dashboard that uses checkmarks to represent completed tasks.
// AI Prompts: Generate C# code that imports a bool[,] into an Aspose.Cells worksheet and applies IsCheckBoxStyle to show checkmarks for true values. | Explain why ImportTwoDimensionArray requires an object[,] and provide the most efficient conversion from bool[,] in C#. | Suggest alternative methods to apply checkbox styling to a range after importing boolean data with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to convert a bool[,] to an object[,] for ImportTwoDimensionArray, load the data into a worksheet starting at A1, apply the IsCheckBoxStyle property to each cell so that true entries appear as checkmarks, and save the workbook as BooleanCheckBox.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define a two‑dimensional array of booleans
        bool[,] boolData = new bool[,]
        {
            { true,  false, true  },
            { false, false, true  },
            { true,  true,  false }
        };

        int rows = boolData.GetLength(0);
        int cols = boolData.GetLength(1);

        // Convert the bool[,] to object[,] because ImportTwoDimensionArray expects an object array
        object[,] objData = new object[rows, cols];
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                objData[i, j] = boolData[i, j];
            }
        }

        // Import the two‑dimensional array into the worksheet starting at cell A1 (row 0, column 0)
        cells.ImportTwoDimensionArray(objData, 0, 0);

        // Apply checkbox style to each imported cell
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                Cell cell = cells[i, j];
                cell.IsCheckBoxStyle = true;   // display as a check box
                // The boolean value is already imported; no need to call PutValue again
            }
        }

        // Save the workbook to a file
        workbook.Save("BooleanCheckBox.xlsx");
    }
}
