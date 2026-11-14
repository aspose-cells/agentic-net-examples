// Title: How to set the width of column 3 using Aspose.Cells Cells.SetColumnWidth in C#
// Common Searches: C# Aspose.Cells set column width by column index example | How to change width of the third column in an Excel file using Aspose.Cells | Set column width to 20 characters in a workbook with Aspose.Cells for .NET
// Tags: Aspose.Cells SetColumnWidth method usage | adjust column width via Cells API | modify column width in .xlsx workbook C# | specify column width in characters Aspose.Cells | set column width by zero‑based index

using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, and uses sheet.Cells.SetColumnWidth(2, 20) to set the third column's width to 20 characters before saving the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the width of column 3 (zero‑based index 2) to 20 characters
        sheet.Cells.SetColumnWidth(2, 20); // column index, width

        // Save the workbook
        workbook.Save("Result.xlsx"); // save
    }
}
