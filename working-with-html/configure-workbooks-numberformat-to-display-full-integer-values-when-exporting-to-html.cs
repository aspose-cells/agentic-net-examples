// Title: Force integer display (no scientific notation) when saving a workbook to HTML with Aspose.Cells for .NET
// Description: Shows how to apply the built‑in integer format (Number = 1, pattern "0") to a cell range via Style and StyleFlag, then export the workbook as HTML so large numbers are rendered in full without scientific notation.
// Keywords: Aspose.Cells .NET | C# HTML export | NumberFormat integer | prevent scientific notation | StyleFlag number format | built‑in format 0 | cell styling for HTML | preserve numeric display | export Excel to HTML | global developer guide
// Common Searches: Aspose.Cells set integer format before HTML export | how to stop scientific notation in HTML output Aspose.Cells | apply built‑in number format 0 to range C# | StyleFlag only number format Aspose.Cells example | save workbook as HTML with full integer values
// Developer Intent: Apply a plain integer NumberFormat to specific cells so the HTML output shows the complete numeric value.
// Use Cases: Web reports that must display exact ID or account numbers without scientific notation. | Publishing Excel‑based dashboards to HTML while keeping totals and counts as whole numbers. | Generating product‑code tables for e‑commerce sites where codes are large integers.
// AI Prompts: Provide C# code that sets the integer NumberFormat ("0") on a range using Aspose.Cells and then saves the workbook as HTML. | Explain how StyleFlag can be used to modify only the number format of cells in Aspose.Cells for .NET. | Show how HtmlSaveOptions respects cell NumberFormat to prevent scientific notation in the generated HTML.

using System;
using Aspose.Cells;

// Shows how to apply the built‑in integer format (Number = 1, pattern "0") to a cell range via Style and StyleFlag, then export the workbook as HTML so large numbers are rendered in full without scientific notation.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells with integer values that could be displayed in scientific notation by default
        worksheet.Cells["A1"].PutValue(123456789);
        worksheet.Cells["A2"].PutValue(9876543210L);

        // Create a style that forces integer display (no decimal places)
        Style integerStyle = workbook.CreateStyle();
        integerStyle.Number = 1; // Built‑in format "0" (plain integer)

        // Apply the style only to the number format of the target range
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;
        worksheet.Cells.CreateRange("A1:A2").ApplyStyle(integerStyle, flag);

        // Save the workbook as HTML
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        workbook.Save("IntegerValues.html", htmlOptions);
    }
}
