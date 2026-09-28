// Title: Apply a light gray background to rows C5:C9 across all columns with Aspose.Cells in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to create a range covering rows 5 to 9 (C5:C9) across the whole worksheet and apply a solid light gray fill. | Show how to define a style with a light gray background and apply it to a multi‑row range in a new workbook using Aspose.Cells for .NET. | Provide a step‑by‑step example of creating a range for rows C5‑C9, setting the ForegroundColor to LightGray, and saving the workbook.
// Common Searches: Aspose.Cells C# how to set background color for rows C5 to C9 | Create a range that spans rows 5‑9 across all columns and apply a solid fill with Aspose.Cells | C# Aspose.Cells style entire row range and save workbook | Apply light gray fill to multiple rows using Aspose.Cells .NET API
// Tags: Aspose.Cells apply style to row range | Aspose.Cells create range across all columns | Aspose.Cells set background color light gray | C# Aspose.Cells format rows C5 to C9 | Aspose.Cells style entire worksheet rows

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, determines the total column count, defines a range that covers rows 5 through 9 across all columns, builds a style with a solid light gray background, applies the style to the entire range, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the total number of columns in the sheet
            int totalColumns = sheet.Cells.MaxColumn + 1; // MaxColumn is zero‑based
            if (totalColumns == 0)
                totalColumns = 1; // Ensure at least one column exists for the range

            // Create a range that covers rows 5 to 9 (zero‑based rows 4‑8)
            // firstRow = 4, firstColumn = 0, totalRows = 5, totalColumns = all columns in the sheet
            Aspose.Cells.Range range = sheet.Cells.CreateRange(4, 0, 5, totalColumns);

            // Create a style with light gray background
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the whole range
            StyleFlag flag = new StyleFlag() { All = true };
            range.ApplyStyle(style, flag);

            // Save the workbook (lifecycle save)
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
