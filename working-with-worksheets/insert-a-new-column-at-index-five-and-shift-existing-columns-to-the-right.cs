// Title: How to insert a column at index 5 and shift existing columns right using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells insert column at position 5 and move other columns | Insert a new column before column F with Aspose.Cells .NET | How to add a column and shift existing columns in an Excel worksheet using Aspose.Cells | Aspose.Cells Cells.InsertColumns example for inserting a single column | Shift columns right after inserting a column in a workbook using Aspose.Cells C#
// Tags: Aspose.Cells Cells.InsertColumns method | insert column at specific index C# | shift worksheet columns right .NET | add column before column F Excel | modify worksheet column layout Aspose.Cells

using Aspose.Cells;

// The sample loads 'input.xlsx', inserts one column at zero‑based index 5 on the first worksheet with Cells.InsertColumns (shifting existing columns to the right), and saves the result as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a new column at index 5 (zero‑based, i.e., before column F)
        // The second parameter (1) specifies that one column is inserted
        worksheet.Cells.InsertColumns(5, 1);

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
