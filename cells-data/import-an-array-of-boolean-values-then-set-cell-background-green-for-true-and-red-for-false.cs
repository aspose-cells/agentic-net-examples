// Title: Import a Boolean array into Excel with Aspose.Cells for .NET and color true cells green, false cells red
// AI Prompts: Load a bool[] into a worksheet starting at cell A1 using Cells.ImportObjectArray, then create a light‑green style for true values. | Create a solid light‑coral style for false values and assign it to each cell whose BoolValue is false. | Save the workbook as an .xlsx file after applying the green and red background styles to the imported Boolean column.
// Common Searches: how to import a bool array into Excel using Aspose.Cells C# | set cell background color based on boolean value with Aspose.Cells .NET | apply green fill to true cells and red fill to false cells in Aspose.Cells | conditional formatting of imported boolean data in Aspose.Cells C# | style cells after ImportObjectArray Aspose.Cells example
// Tags: boolean data import Aspose.Cells | conditional cell fill Aspose.Cells | true false cell styling C# | green red background Aspose.Cells | save styled workbook .xlsx

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new workbook, converts a bool[] to an object array, imports it vertically into column A, defines solid light‑green and light‑coral styles, applies the appropriate style to each cell according to its Boolean value, and saves the result as BooleanArrayStyled.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Boolean array to import
        bool[] boolArray = new bool[] { true, false, true, true, false };

        // Convert bool[] to object[] because ImportObjectArray expects object[]
        object[] objArray = Array.ConvertAll(boolArray, b => (object)b);

        // Import the boolean values vertically starting at cell A1 (row 0, column 0)
        cells.ImportObjectArray(objArray, 0, 0, true);

        // Create style for true values (green background)
        Style trueStyle = workbook.CreateStyle();
        trueStyle.ForegroundColor = Color.LightGreen;
        trueStyle.Pattern = BackgroundType.Solid;

        // Create style for false values (red background)
        Style falseStyle = workbook.CreateStyle();
        falseStyle.ForegroundColor = Color.LightCoral;
        falseStyle.Pattern = BackgroundType.Solid;

        // Apply the appropriate style to each imported cell based on its boolean value
        for (int i = 0; i < boolArray.Length; i++)
        {
            Cell cell = cells[i, 0]; // Cells are imported in column A
            if (cell.BoolValue)
                cell.SetStyle(trueStyle);
            else
                cell.SetStyle(falseStyle);
        }

        // Save the workbook
        workbook.Save("BooleanArrayStyled.xlsx");
    }
}
