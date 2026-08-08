// Title: C# – Import Boolean Array into Aspose.Cells and color cells green for true, red for false
// Description: Creates a Workbook, imports a bool[] into column A with ImportObjectArray, defines solid light‑green and light‑coral styles, applies the appropriate style to each cell based on its BoolValue, and saves the file as BooleanArrayStyled.xlsx.
// Keywords: Aspose.Cells | C# | ImportObjectArray | boolean array | cell background color | conditional styling | green true red false | Excel export
// Common Searches: Aspose.Cells import bool array C# | set cell color based on boolean value Aspose.Cells | ImportObjectArray conditional formatting example | color true false cells in Excel using Aspose | C# Aspose.Cells style cells programmatically
// Developer Intent: Load a bool[] into a worksheet and automatically color true cells green and false cells red.
// Use Cases: Status dashboards that highlight pass/fail flags with green and red cells. | Checklists where completed items (true) appear in green and pending items (false) in red. | Automated test result reports that visually separate passed and failed cases.
// AI Prompts: Generate C# code using Aspose.Cells to import a bool[] into a worksheet and apply a light‑green background for true cells and a light‑coral background for false cells. | Show how to combine ImportObjectArray with a loop that sets a style based on each cell's BoolValue in Aspose.Cells. | Explain step‑by‑step how to create two solid background styles and assign them to cells after importing a boolean array with Aspose.Cells.

using Aspose.Cells;
using System;
using System.Drawing;

// Creates a Workbook, imports a bool[] into column A with ImportObjectArray, defines solid light‑green and light‑coral styles, applies the appropriate style to each cell based on its BoolValue, and saves the file as BooleanArrayStyled.xlsx.
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

        // Convert the bool array to an object array for ImportObjectArray
        object[] objArray = Array.ConvertAll(boolArray, b => (object)b);

        // Import the boolean values vertically starting at cell A1 (row 0, column 0)
        cells.ImportObjectArray(objArray, 0, 0, true);

        // Create a style with green background for true values
        Style trueStyle = workbook.CreateStyle();
        trueStyle.ForegroundColor = Color.LightGreen;
        trueStyle.Pattern = BackgroundType.Solid;

        // Create a style with red background for false values
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

        // Save the workbook to a file
        workbook.Save("BooleanArrayStyled.xlsx");
    }
}
