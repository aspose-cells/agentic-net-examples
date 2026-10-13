// Title: Create a custom number format in Aspose.Cells for .NET that shows negative values in red parentheses
// AI Prompts: Generate C# code that defines a cell style with the format "#,##0.00;[Red](#,##0.00)" and applies it to a worksheet using Aspose.Cells. | Write a script that inserts a negative number into an Excel cell and formats it to appear in red parentheses via Aspose.Cells custom number format. | Produce a complete Aspose.Cells example that creates a workbook, sets a red‑parentheses custom format for negatives, and saves the file.
// Common Searches: Aspose.Cells C# custom number format for red parentheses on negative numbers | How to apply red parentheses style to negative values in an Excel file using Aspose.Cells | C# example of setting '#,##0.00;[Red](#,##0.00)' format with Aspose.Cells | Formatting negative numbers with color and parentheses in Aspose.Cells .NET
// Tags: custom number format negative red parentheses Aspose.Cells | cell style custom format C# Excel | Aspose.Cells set custom number format string | negative value formatting red parentheses .NET | save workbook with styled cells Aspose.Cells

using Aspose.Cells;

// The program creates a new workbook, writes a negative value to cell A1, applies the custom number format "#,##0.00;[Red](#,##0.00)" so negatives appear in red parentheses, and saves the file as CustomNumberFormat.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a sample negative value into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue(-1234.56);

        // Define a custom number format:
        // Positive numbers: normal display
        // Negative numbers: red color and enclosed in parentheses
        // Format string: "#,##0.00;[Red](#,##0.00)"
        Style style = cell.GetStyle();
        style.Custom = "#,##0.00;[Red](#,##0.00)";
        cell.SetStyle(style);

        // Save the workbook (lifecycle save)
        workbook.Save("CustomNumberFormat.xlsx");
    }
}
