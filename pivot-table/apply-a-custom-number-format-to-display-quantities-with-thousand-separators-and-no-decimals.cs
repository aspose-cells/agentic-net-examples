// Title: Apply a custom number format with thousand separators and no decimals to a cell range using Aspose.Cells for .NET
// Common Searches: how to format numbers with commas in Aspose.Cells C# | apply #,##0 number format to a range using Aspose.Cells | set cell style to show thousand separators without decimals in .NET
// Tags: custom number format #,##0 Aspose.Cells | StyleFlag NumberFormat property Aspose.Cells | apply thousand separator style to cell range .NET | Excel numeric style without decimals Aspose | create integer display style Aspose.Cells

using Aspose.Cells;

// The example creates a workbook, writes integer quantities to cells A1‑A3, defines a style with the custom format "#,##0" to display thousand separators and no decimal places, applies this style to the range using a StyleFlag that enables only number formatting, and saves the file as Quantities.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Sample quantity values
        sheet.Cells["A1"].PutValue(1234);
        sheet.Cells["A2"].PutValue(56789);
        sheet.Cells["A3"].PutValue(1234567);

        // Create a style with a custom number format that adds thousand separators and no decimals
        Style qtyStyle = workbook.CreateStyle();
        qtyStyle.Custom = "#,##0";

        // Define which style attributes to apply (only number format)
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the style to the range containing the quantities
        sheet.Cells.CreateRange("A1:A3").ApplyStyle(qtyStyle, flag);

        // Save the workbook (lifecycle rule)
        workbook.Save("Quantities.xlsx");
    }
}
