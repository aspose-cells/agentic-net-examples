// Title: C# – Create a non‑contiguous named range with UnionRange and apply a custom style using Aspose.Cells
// Description: Demonstrates how to build a UnionRange that covers cells A1:A5 and C1:C5, assign it the name "MyUnionRange", populate the cells, define a light‑green solid background with a bold dark‑blue font, apply the style to the entire range, and save the workbook as NamedUnionRange.xlsx.
// Keywords: Aspose.Cells | C# | .NET | UnionRange | named range | non‑contiguous cells | custom style | background color | bold font | Excel automation
// Common Searches: Aspose.Cells create UnionRange non‑adjacent cells | named range spanning multiple columns Aspose.Cells C# | apply custom style to UnionRange Aspose.Cells | set background color for non‑contiguous range Aspose.Cells | how to name a UnionRange in Aspose.Cells
// Developer Intent: Define a UnionRange that includes non‑adjacent cells, give it a name, and format it with a custom style in a .NET workbook.
// Use Cases: Highlight header rows in separate columns of a financial report with a unified style. | Create a dashboard template where specific sections in columns A and C are visually distinguished. | Generate printable Excel forms that require styled, non‑contiguous blocks for branding or section titles.
// AI Prompts: Show C# code to create a UnionRange for A1:A5 and C1:C5, name it, and apply a solid light‑green background with bold dark‑blue text using Aspose.Cells. | Explain how to modify the style of an existing UnionRange after the workbook has been saved. | Provide a step‑by‑step guide to assign a custom style to a named non‑contiguous range in Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;

// Demonstrates how to build a UnionRange that covers cells A1:A5 and C1:C5, assign it the name "MyUnionRange", populate the cells, define a light‑green solid background with a bold dark‑blue font, apply the style to the entire range, and save the workbook as NamedUnionRange.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a union range that covers non‑contiguous cells A1:A5 and C1:C5
        UnionRange unionRange = workbook.Worksheets.CreateUnionRange("A1:A5,C1:C5", 0);

        // Assign a name to the union range
        unionRange.Name = "MyUnionRange";

        // Populate the cells so the result can be seen
        for (int i = 0; i < 5; i++)
        {
            worksheet.Cells[i, 0].PutValue($"A{i + 1}");
            worksheet.Cells[i, 2].PutValue($"C{i + 1}");
        }

        // Create a custom style
        Style style = workbook.CreateStyle();
        style.ForegroundColor = Color.LightGreen;
        style.Pattern = BackgroundType.Solid;
        style.Font.IsBold = true;
        style.Font.Color = Color.DarkBlue;

        // Apply the style to the entire union range
        unionRange.ApplyStyle(style, new StyleFlag { All = true });

        // Save the workbook
        workbook.Save("NamedUnionRange.xlsx");
    }
}
