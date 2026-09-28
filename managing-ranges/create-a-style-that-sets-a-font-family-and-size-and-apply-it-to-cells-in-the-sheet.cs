// Title: Create an Arial 12 font style and apply it to a cell range (A1:C3) using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Style with Font.Name = "Arial" and Font.Size = 12, then applies it to the range A1:C3 in an Aspose.Cells workbook using a StyleFlag that limits the change to font attributes. | Show how to define a reusable style in Aspose.Cells, set additional font properties such as Bold and Italic, and apply the style to a specific range while preserving other cell formatting. | Generate a complete .NET console program that builds a new workbook, creates a custom font style, applies it to a defined range, and saves the file as StyledWorkbook.xlsx.
// Common Searches: Aspose.Cells C# apply custom font style to multiple cells | How to use StyleFlag to change only font in Aspose.Cells range | C# example for creating and applying a style to A1:C3 with Aspose.Cells | Set Arial 12 font for a range in an Aspose.Cells workbook programmatically | Apply bold and italic font to a cell block using Aspose.Cells .NET
// Tags: Aspose.Cells create custom font style C# | apply style to cell range Aspose.Cells | StyleFlag font-only application Aspose.Cells | set Arial 12 font Aspose.Cells workbook | define and reuse style Aspose.Cells .NET | range A1:C3 style application Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates creating a new Workbook, defining a Style with Arial 12 font, using a StyleFlag to apply only the font attributes to the range A1:C3, and saving the workbook as StyledWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a new style object
            Style style = workbook.CreateStyle();

            // Set the desired font family and size
            style.Font.Name = "Arial";
            style.Font.Size = 12;

            // Define a range to which the style will be applied (e.g., A1:C3)
            // Use fully qualified name to avoid conflict with System.Range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1", "C3");

            // Apply the style to the range (only font attributes)
            StyleFlag flag = new StyleFlag
            {
                Font = true
            };
            range.ApplyStyle(style, flag);

            // Define output file path
            string outputPath = "StyledWorkbook.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
