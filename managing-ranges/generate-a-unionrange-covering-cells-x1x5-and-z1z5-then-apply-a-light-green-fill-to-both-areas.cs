// Title: Apply a light green fill to a union range covering X1:X5 and Z1:Z5 with Aspose.Cells for .NET
// AI Prompts: Create a union range for cells X1:X5 and Z1:Z5 and set a solid light‑green background style using Aspose.Cells in C#. | Use Aspose.Cells to apply a light green fill to a non‑adjacent range defined by X1:X5,Z1:Z5 and save the workbook.
// Common Searches: how to set background color for multiple non‑adjacent cells in Aspose.Cells C# | Aspose.Cells union range X1:X5 Z1:Z5 fill color example | C# code to apply solid light green style to a comma‑separated cell range in Aspose.Cells | create and style a union range with Aspose.Cells workbook | apply style to noncontiguous cells using Aspose.Cells CreateRange method
// Tags: union range styling Aspose.Cells C# | noncontiguous cell fill color Aspose.Cells | light green fill style Aspose.Cells | CreateRange with comma separated addresses Aspose.Cells | StyleFlag All property Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The program creates a new workbook, defines a union range covering X1:X5 and Z1:Z5, applies a solid light‑green fill style to the range, and saves the file as UnionRangeLightGreen.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Create a union range that includes cells X1:X5 and Z1:Z5
            var unionRange = worksheet.Cells.CreateRange("X1:X5,Z1:Z5");

            // Define a style with a light green solid fill
            var style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightGreen;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the entire union range
            var flag = new StyleFlag() { All = true };
            unionRange.ApplyStyle(style, flag);

            // Save the workbook to a file
            workbook.Save("UnionRangeLightGreen.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
