// Title: Merge cells A1:F2 into a title block using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to merge the range A1:F2 on the first worksheet and write a title string into the merged cell. | Create a title block that spans rows 1‑2 and columns A‑F, then save the workbook as an .xlsx file with Aspose.Cells. | Programmatically merge multiple rows and columns into a single cell region and assign a value using the Aspose.Cells C# API.
// Common Searches: aspnet c# merge cells A1 to F2 Aspose.Cells example | how to create a title block in Excel using Aspose.Cells for .NET | Aspose.Cells merge range across rows and columns C# | set value in merged Excel cells with Aspose.Cells C# | save workbook after merging cells Aspose.Cells .NET
// Tags: Aspose.Cells merge cell range C# | Excel title block Aspose.Cells | set merged cell value Aspose.Cells | save workbook Aspose.Cells .xlsx | merge multiple rows columns Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, merges cells A1 through F2 into a single range, writes "Report Title" into the merged cell, and saves the file as TitleBlock.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells from A1 to F2 (rows 0-1, columns 0-5)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        sheet.Cells.Merge(0, 0, 2, 6);

        // Optional: set a title in the merged region
        sheet.Cells["A1"].PutValue("Report Title");

        // Save the workbook to a file
        workbook.Save("TitleBlock.xlsx");
    }
}
