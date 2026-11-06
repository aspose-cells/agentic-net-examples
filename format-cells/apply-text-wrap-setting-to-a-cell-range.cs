// Title: How to wrap text in a specific cell range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells wrap text for cells A1:C3 | apply text wrapping to a range of cells using Aspose.Cells .NET | how to set IsTextWrapped on multiple cells with Aspose.Cells StyleFlag | Aspose.Cells example to enable wrap text in a worksheet range | using StyleFlag to apply wrap text style to a cell range in C#
// Tags: text wrap style application Aspose.Cells | StyleFlag All usage Aspose.Cells | define cell range A1:C3 Aspose.Cells | save workbook with text wrapping Aspose.Cells | C# Aspose.Cells wrap text example

using Aspose.Cells;
using System;

// The C# program creates a new Workbook, accesses the first worksheet, defines the range A1:C3, creates a Style with IsTextWrapped set to true, uses a StyleFlag with All enabled, applies the style to the range, and saves the workbook as TextWrapExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the cell range to apply text wrap (e.g., A1:C3)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:C3");

            // Create a style with text wrap enabled
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;

            // Define a StyleFlag to apply all style attributes
            StyleFlag flag = new StyleFlag();
            flag.All = true;

            // Apply the wrap style to the specified range
            range.ApplyStyle(wrapStyle, flag);

            // Save the workbook (lifecycle save)
            string outputPath = "TextWrapExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
