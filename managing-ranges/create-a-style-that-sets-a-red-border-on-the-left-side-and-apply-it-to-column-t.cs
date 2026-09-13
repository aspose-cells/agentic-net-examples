// Title: How to create and apply a red left border style to column T using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that defines a thin red left border style and applies it exclusively to column T. | Show how to use a StyleFlag to limit formatting to borders when applying a custom style to a specific column in an Excel worksheet via Aspose.Cells.
// Common Searches: Aspose.Cells C# set only the left border red for an entire column | Apply a custom border style to column T in an Excel file using Aspose.Cells .NET | How to use StyleFlag to apply border formatting without changing other cell attributes in Aspose.Cells | Create a thin red left border for column T programmatically with Aspose.Cells | Aspose.Cells apply style to column by index 19 in C#
// Tags: red left border style Aspose.Cells C# | column T style application Aspose.Cells | StyleFlag border‑only formatting Aspose.Cells | thin left border Excel column .NET | column formatting using Aspose.Cells C#

using Aspose.Cells;
using System.Drawing;

// The program creates a new workbook, defines a style with a thin red left border, uses a StyleFlag to restrict formatting to borders only, applies the style to column T (index 19), and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Create a style with a red left border
        Style leftRedBorderStyle = workbook.CreateStyle();
        leftRedBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
        leftRedBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Red;

        // Specify that only border settings should be applied
        StyleFlag flag = new StyleFlag();
        flag.Borders = true;

        // Apply the style to column T (zero‑based index 19)
        sheet.Cells.Columns[19].ApplyStyle(leftRedBorderStyle, flag);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
