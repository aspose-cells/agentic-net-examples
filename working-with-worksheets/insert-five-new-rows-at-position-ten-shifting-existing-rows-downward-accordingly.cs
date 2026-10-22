// Title: How to insert five rows at row 10 in the first worksheet of an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Add five new rows starting at row 10 of the first worksheet, shift existing rows down, and save the workbook using Aspose.Cells in C#. | Refactor the sample to accept a row index and row count as parameters, insert those rows into any worksheet, and write the updated file to a specified path.
// Common Searches: Aspose.Cells C# insert rows at row 10 and shift down existing rows | C# code to add multiple rows to the first sheet of an Excel workbook using Aspose.Cells | How to programmatically insert five rows at a specific position in an .xlsx file with Aspose.Cells | Insert rows in Excel worksheet using Aspose.Cells .NET API example | Insert rows at index 9 using Aspose.Cells Cells.InsertRows method
// Tags: cells.insertrows method Aspose.Cells | insert multiple rows Excel .NET | add rows to first worksheet Aspose.Cells | shift rows down Excel workbook C# | dynamic row insertion Aspose.Cells

using Aspose.Cells;

// Loads an existing .xlsx workbook, inserts five rows at the 10th row of the first worksheet (shifting existing rows downward) using Aspose.Cells for .NET, and saves the result to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert 5 rows at position 10 (zero‑based index 9)
        sheet.Cells.InsertRows(9, 5);

        // Save the modified workbook (replace with your target file)
        workbook.Save("output.xlsx");
    }
}
