// Title: Import GUID List into Excel and Convert to Uppercase with Aspose.Cells for .NET
// Description: Shows how to generate a collection of GUID strings, import them vertically into a worksheet using Cells.ImportArrayList, convert each value to uppercase, and save the workbook as an .xlsx file with Aspose.Cells for .NET.
// Keywords: ImportArrayList | GUID strings | uppercase cells | Aspose.Cells | C# Excel automation | vertical data import | ArrayList to worksheet | cell value transformation | unique identifier formatting | Excel file generation
// Common Searches: Aspose.Cells ImportArrayList example | import GUID list into Excel C# | convert cell text to uppercase Aspose.Cells | vertical import of strings Aspose.Cells | C# Aspose.Cells write GUIDs to worksheet
// Developer Intent: Add a collection of GUID strings to a worksheet and display them in uppercase using Aspose.Cells.
// Use Cases: Create a tracking sheet where each row contains a unique identifier in a standardized uppercase format. | Generate compliance reports that require GUIDs to be presented in uppercase for readability. | Build a template that automatically populates a column with GUIDs and enforces consistent casing before distribution.
// AI Prompts: Write C# code that uses Aspose.Cells to import an ArrayList of GUID strings into column A and convert the values to uppercase. | Explain how to apply a style or transformation in Aspose.Cells to display imported GUIDs in uppercase without iterating over each cell. | Show an alternative Aspose.Cells API call that imports a vertical list of strings and forces uppercase formatting during import.

using System;
using System.Collections;
using Aspose.Cells;

// Shows how to generate a collection of GUID strings, import them vertically into a worksheet using Cells.ImportArrayList, convert each value to uppercase, and save the workbook as an .xlsx file with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Build an ArrayList containing GUID strings
        ArrayList guidList = new ArrayList();
        for (int i = 0; i < 5; i++)
        {
            guidList.Add(Guid.NewGuid().ToString());
        }

        // Import the GUID list vertically starting at cell A1 (row 0, column 0)
        // Parameters: (ArrayList, firstRow, firstColumn, isVertical)
        cells.ImportArrayList(guidList, 0, 0, true);

        // Convert each imported GUID to uppercase
        for (int row = 0; row < guidList.Count; row++)
        {
            Cell cell = cells[row, 0];
            if (cell.Value != null)
            {
                // Overwrite the cell value with its uppercase representation
                cell.PutValue(cell.StringValue.ToUpper());
            }
        }

        // Save the workbook to a file
        workbook.Save("GuidUppercase.xlsx");
    }
}
