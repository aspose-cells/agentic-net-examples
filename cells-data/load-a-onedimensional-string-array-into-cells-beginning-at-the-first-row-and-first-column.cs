// Title: Import a 1‑D string array into cell A1 with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, converts a string[] to object[] and uses Worksheet.Cells.ImportObjectArray to write the values horizontally starting at row 0, column 0 (cell A1), then saves the file as OneDimensionalStringArray.xlsx.
// Keywords: Aspose.Cells C# import string array | ImportObjectArray example | load 1‑D array into Excel | horizontal array import Aspose | convert string[] to object[]
// Common Searches: Aspose.Cells import string array C# | ImportObjectArray horizontal example | write string[] to Excel with Aspose | populate first row from array Aspose.Cells
// Developer Intent: Insert a one‑dimensional string array into the first row and first column of an Excel worksheet using Aspose.Cells.
// Use Cases: Create a header row from a list of column names stored in a string array. | Place a series of labels across the top of a report sheet. | Initialize a simple data table where the first row is populated from predefined strings.
// AI Prompts: Generate C# code that uses Aspose.Cells to import a string[] into cell A1 horizontally. | Show how to import a one‑dimensional array vertically with ImportObjectArray in Aspose.Cells. | Explain how to convert any primitive array (int[], double[], etc.) to object[] for ImportObjectArray.

using System;
using Aspose.Cells;

// Creates a new Workbook, converts a string[] to object[] and uses Worksheet.Cells.ImportObjectArray to write the values horizontally starting at row 0, column 0 (cell A1), then saves the file as OneDimensionalStringArray.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // One‑dimensional string array to be loaded into the sheet
        string[] stringArray = new string[] { "Alpha", "Beta", "Gamma", "Delta" };

        // Convert the string array to an object array because ImportObjectArray expects object[]
        object[] objArray = Array.ConvertAll(stringArray, s => (object)s);

        // Import the array horizontally (isVertical = false) starting at the first row (0) and first column (0)
        worksheet.Cells.ImportObjectArray(objArray, 0, 0, false);

        // Save the workbook to a file
        workbook.Save("OneDimensionalStringArray.xlsx");
    }
}
