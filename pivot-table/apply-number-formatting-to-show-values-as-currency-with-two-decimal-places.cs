// Title: Apply a custom currency number format ($#,##0.00) with two decimal places to an Excel cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# format cell as currency with two decimal places | set custom number format $#,##0.00 in Excel using Aspose.Cells .NET | apply currency style to a worksheet cell programmatically with Aspose.Cells | C# Aspose.Cells example for currency formatting and saving workbook | how to use a custom style to display numbers as currency in Aspose.Cells
// Tags: Aspose.Cells custom currency format | C# apply number format to Excel cell | Excel cell style currency two decimals Aspose | Aspose.Cells set custom style workbook | currency number format $#,##0.00 Aspose.Cells

using Aspose.Cells;

// The sample creates a new workbook, writes the numeric value 1234.5 to cell A1, applies the custom number format "$#,##0.00" to display the value as currency with two decimal places, and saves the file as FormattedCurrency.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Example: put a numeric value into cell A1
        var cell = sheet.Cells["A1"];
        cell.PutValue(1234.5);

        // Get the current style of the cell
        var style = cell.GetStyle();

        // Apply a custom number format for currency with two decimal places
        // Format: $#,##0.00  (e.g., $1,234.50)
        style.Custom = "$#,##0.00";

        // Assign the modified style back to the cell
        cell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("FormattedCurrency.xlsx");
    }
}
