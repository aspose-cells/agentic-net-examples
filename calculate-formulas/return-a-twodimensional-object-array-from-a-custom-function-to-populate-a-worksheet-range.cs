// Title: How to populate an Excel worksheet range from a two‑dimensional object[,] array using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# fill worksheet range with object[,] data | set Excel cell values from a 2D object array using Aspose.Cells | populate Excel sheet starting at A1 from custom function returning object[,] in C# | create range and assign multi‑dimensional array Aspose.Cells .NET | write 3x3 data array to Excel workbook using Aspose.Cells C#
// Tags: Worksheet.Cells.CreateRange object array | assign 2D object[,] to Excel range | Aspose.Cells write multi‑dimensional data to .xlsx | C# load custom data array into worksheet | set range value using object[,] Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, obtains a 3x3 two‑dimensional object[,] array from a custom function, determines its dimensions, and uses Worksheet.Cells.CreateRange to assign the array directly to the range beginning at A1, then saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve a two‑dimensional object array from a custom function
        object[,] data = GetSampleData();

        // Determine the size of the array
        int rows = data.GetLength(0);
        int cols = data.GetLength(1);

        // Populate the worksheet range starting at cell A1 with the array
        // CreateRange(row, column, totalRows, totalColumns) sets the target range
        sheet.Cells.CreateRange(0, 0, rows, cols).Value = data;

        // Save the workbook (lifecycle rule)
        workbook.Save("Output.xlsx");
    }

    // Custom function that returns a two‑dimensional object array
    static object[,] GetSampleData()
    {
        // Example 3x3 data set
        object[,] arr = new object[3, 3];
        arr[0, 0] = "Name";
        arr[0, 1] = "Age";
        arr[0, 2] = "Score";

        arr[1, 0] = "Alice";
        arr[1, 1] = 30;
        arr[1, 2] = 85.5;

        arr[2, 0] = "Bob";
        arr[2, 1] = 28;
        arr[2, 2] = 92.0;

        return arr;
    }
}
