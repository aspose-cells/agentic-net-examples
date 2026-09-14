// Title: How to insert a column at index 5 and shift existing columns right using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a single column at zero‑based index 5 in an existing Excel workbook with Aspose.Cells for .NET and save the file. | Write C# code that adds three consecutive columns starting at column F while preserving formulas and formatting using Aspose.Cells. | Create a reusable C# method that accepts a worksheet and a column index, inserts a column at that position, and returns the updated workbook.
// Common Searches: Aspose.Cells C# insert column at index 5 before column F | how to shift worksheet columns right after inserting a column with Aspose.Cells .NET | insert multiple columns at a specific position in Excel using Aspose.Cells C# example
// Tags: Aspose.Cells InsertColumns method C# | add column to Excel worksheet .NET | shift worksheet columns right Aspose.Cells | insert column before column F Excel C# | modify worksheet structure programmatically Aspose.Cells

using Aspose.Cells;

// The example loads an existing Excel workbook, accesses the first worksheet, inserts one column at zero‑based index 5 (before column F) which shifts all subsequent columns to the right, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new column at index 5 (zero‑based, i.e., before column F)
        // The second parameter (1) specifies the number of columns to insert.
        sheet.Cells.InsertColumns(5, 1);

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
