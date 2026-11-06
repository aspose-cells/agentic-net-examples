// Title: How to apply an accounting number format to a total sales cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# apply accounting number format to a specific cell | Set accounting style for total sales value in Excel with Aspose.Cells .NET | Use custom number format "$#,##0.00_);[Red]($#,##0.00)" in Aspose.Cells C# | Formatting a cell as accounting currency with negative numbers in red using Aspose.Cells | Apply accounting number format to a workbook cell in C#
// Tags: accounting number format Aspose.Cells | custom number format C# | format cell as accounting Excel | apply style to workbook cell .NET | currency accounting format Aspose.Cells

using Aspose.Cells;

// The example creates a new workbook, writes a total sales amount to cell B2, applies a custom accounting number format "$#,##0.00_);[Red]($#,##0.00)" to that cell, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a sample total sales value into cell B2
        Cell totalSalesCell = sheet.Cells["B2"];
        totalSalesCell.PutValue(12345.67);

        // Apply accounting number format to the cell
        // Accounting format example: "$#,##0.00_);[Red]($#,##0.00)"
        Style style = totalSalesCell.GetStyle();
        style.Custom = "$#,##0.00_);[Red]($#,##0.00)";
        totalSalesCell.SetStyle(style);

        // Save the workbook (lifecycle rule: save)
        workbook.Save("Output.xlsx");
    }
}
