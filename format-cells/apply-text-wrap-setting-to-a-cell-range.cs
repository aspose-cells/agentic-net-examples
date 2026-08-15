// Title: Apply Text Wrap to a Cell Range with Aspose.Cells for .NET
// Description: Creates a workbook, fills a 3×3 range with long text, defines a style with IsTextWrapped = true, uses a StyleFlag to apply only the wrap‑text attribute, applies the style to range A1:C3, auto‑fits rows, and saves the file as an XLSX document.
// Keywords: Aspose.Cells wrap text C# | apply text wrap range Aspose.Cells | StyleFlag wrap text .NET | auto fit rows after wrap Aspose | IsTextWrapped property example
// Common Searches: how to enable text wrap for a range in Aspose.Cells .NET | apply wrap text style to multiple cells using StyleFlag | auto fit rows after wrapping text Aspose.Cells | set IsTextWrapped for a cell range C#
// Developer Intent: Enable text wrapping for a specific cell range and adjust row heights automatically.
// Use Cases: Display lengthy product descriptions in invoice tables without expanding column width. | Format multi‑line headers in dashboard worksheets for clearer presentation. | Generate reports where comments or notes need to wrap within a defined block of cells.
// AI Prompts: Show C# code to apply text wrap to a dynamic range based on content length using Aspose.Cells. | Provide an example that toggles wrap text on or off for a selected range and refreshes row heights. | Explain how StyleFlag limits style changes to only the WrapText property in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, fills a 3×3 range with long text, defines a style with IsTextWrapped = true, uses a StyleFlag to apply only the wrap‑text attribute, applies the style to range A1:C3, auto‑fits rows, and saves the file as an XLSX document.
class ApplyWrapTextToRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a 3x3 range with long text that needs wrapping
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue("This is a long text that should wrap inside the cell.");
                }
            }

            // Create a style with text wrapping enabled
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Create a style flag to apply only the WrapText property
            StyleFlag flag = new StyleFlag();
            flag.WrapText = true;

            // Define the range A1:C3 and apply the wrap style using the flag
            Aspose.Cells.Range range = cells.CreateRange(0, 0, 3, 3);
            range.ApplyStyle(wrapStyle, flag);

            // Auto‑fit rows so the wrapped text becomes visible
            sheet.AutoFitRows();

            // Save the workbook
            workbook.Save("WrapTextRangeDemo.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
