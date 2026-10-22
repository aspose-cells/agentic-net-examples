// Title: Apply accounting number format with red negatives to a specific cell using Aspose.Cells for .NET
// Common Searches: how to apply accounting format with red negative numbers to a cell using Aspose.Cells C# | Aspose.Cells set custom accounting style for Excel cell C10 in .NET | format total sales column as accounting with red negatives in Aspose.Cells | C# Aspose.Cells apply custom number format string to a worksheet cell | save workbook after applying accounting number format using Aspose.Cells
// Tags: accounting style customization Aspose.Cells | C# set cell number format | negative values displayed in red Excel | save workbook after styling Aspose.Cells | worksheet cell formatting .NET

using Aspose.Cells;

// The example loads an existing Excel workbook, accesses the first worksheet, retrieves cell C10, applies a custom accounting number format that shows negative values in red, and saves the workbook to a new file with the formatting applied.
class Program
{
    static void Main()
    {
        // Load the existing workbook (create/load rule)
        Workbook workbook = new Workbook("SalesReport.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Locate the total sales cell (example: cell C10)
        Cell totalSalesCell = sheet.Cells["C10"];

        // Apply accounting number format
        // Format: $#,##0.00_);[Red]($#,##0.00) – standard accounting format with red negative numbers
        Style style = totalSalesCell.GetStyle();
        style.Custom = "$#,##0.00_);[Red]($#,##0.00)";
        totalSalesCell.SetStyle(style);

        // Save the workbook with the applied format (save rule)
        workbook.Save("SalesReport_Formatted.xlsx");
    }
}
