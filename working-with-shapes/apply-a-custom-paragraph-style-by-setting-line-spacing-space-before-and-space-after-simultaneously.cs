// Title: Apply a custom paragraph style with font size, color, text wrap, and attempt line‑spacing settings to a single cell using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook, defines a Style with a 12‑point blue font, enables text wrapping, attempts to set line spacing, space before, and space after, and applies the style to cell A1 using a StyleFlag with All = true. | Explain how Aspose.Cells handles paragraph formatting such as line spacing, space before, and space after, and why these properties are not supported in Excel. | Show how to create a Range for "A1", apply a custom Style with a StyleFlag, and save the workbook as an .xlsx file, including proper error handling.
// Common Searches: Aspose.Cells C# set line spacing for a cell style and why it fails | How to apply font color, size, and text wrap to a specific cell using Aspose.Cells .NET | Using StyleFlag All to apply a custom style to a range in Aspose.Cells | Can Excel paragraph spacing be controlled through Aspose.Cells API
// Tags: apply custom cell style Aspose.Cells C# | set font size color wrap Aspose.Cells | paragraph spacing limitation Excel Aspose.Cells | use StyleFlag All Aspose.Cells | create range and apply style Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes text to cell A1, builds a custom Style with a 12‑point blue font, enables text wrapping, and attempts to set line‑spacing, space before, and space after (which Excel does not support). The Style is applied to the A1 range using a StyleFlag with All = true, and the workbook is saved as CustomParagraphStyle.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Write some text into cell A1.
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Sample paragraph text");

            // Create a custom style.
            Style customStyle = workbook.CreateStyle();

            // Example of style settings (Excel does not support paragraph line spacing).
            customStyle.Font.Size = 12;               // Set font size.
            customStyle.Font.Color = System.Drawing.Color.Blue; // Set font color.
            customStyle.IsTextWrapped = true;         // Enable text wrapping.

            // Define which parts of the style should be applied.
            StyleFlag flag = new StyleFlag
            {
                All = true // Apply all style attributes.
            };

            // Apply the custom style to cell A1 using a fully qualified Range to avoid ambiguity.
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1");
            range.ApplyStyle(customStyle, flag);

            // Save the workbook.
            string outputPath = "CustomParagraphStyle.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
