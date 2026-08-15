// Title: Aspose.Cells .NET: Create UnionRange "M1:M5,O1:O5" and Apply Uniform Number Format
// Description: Shows how to build a UnionRange covering M1:M5 and O1:O5, define a custom numeric style (e.g., two decimals), apply it with a StyleFlag, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | UnionRange | number format | custom numeric style | StyleFlag | C# | .NET | Excel formatting | non‑adjacent cells | M1:M5 | O1:O5
// Common Searches: Aspose.Cells create union range | apply number format to union range Aspose | C# set custom numeric format for specific cells | StyleFlag only number format Aspose.Cells | format non adjacent columns Excel C#
// Developer Intent: Create a UnionRange for M1:M5 and O1:O5 and set the same numeric format for all cells.
// Use Cases: Financial statements where two separate columns need identical decimal precision. | Data export templates that require consistent number formatting across non‑adjacent columns. | Reusable Excel report templates that apply a shared style to multiple cell blocks.
// AI Prompts: Write C# code using Aspose.Cells to create a UnionRange "M1:M5,O1:O5" and apply a two‑decimal number format. | Show how to use StyleFlag to change only the number format of a UnionRange in Aspose.Cells .NET. | Explain how to reuse a Style object for multiple UnionRanges with the same numeric pattern in C#.

using System;
using Aspose.Cells;

// Shows how to build a UnionRange covering M1:M5 and O1:O5, define a custom numeric style (e.g., two decimals), apply it with a StyleFlag, and save the workbook using Aspose.Cells for .NET.
class UnionRangeNumberFormatExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that includes columns M and O rows 1 to 5
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("M1:M5,O1:O5", 0);

        // Define a uniform number format (e.g., two decimal places)
        Style numberStyle = workbook.CreateStyle();
        numberStyle.Custom = "0.00";

        // Specify that only the number format should be applied
        StyleFlag flag = new StyleFlag();
        flag.NumberFormat = true;

        // Apply the number format to the entire union range
        unionRange.ApplyStyle(numberStyle, flag);

        // Save the workbook
        workbook.Save("UnionRangeNumberFormat.xlsx");
    }
}
