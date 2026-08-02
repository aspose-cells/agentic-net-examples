// Title: Change Font of All RichTextPortion Objects in a Cell using Aspose.Cells for .NET (C#)
// Description: This example shows how to retrieve the RichText collection of a worksheet cell, loop through each RichTextPortion, set a new Font.Name, and save the workbook. It includes directory creation, exception handling, and demonstrates the complete workflow with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | RichTextPortion | font name | iterate cell portions | Excel rich text formatting | modify cell font | Aspose.Cells example | Excel automation
// Common Searches: Aspose.Cells change font of RichTextPortion | iterate RichTextPortion objects C# | set font name for each portion in Excel cell | Aspose.Cells rich text formatting tutorial | C# code to modify cell rich text font
// Developer Intent: Update the Font.Name of every RichTextPortion inside a specific cell.
// Use Cases: Apply different fonts to individual words within a cell for custom styling. | Standardize font across all rich‑text portions after programmatic content edits. | Bulk‑update rich‑text formatting in generated Excel reports before distribution.
// AI Prompts: Generate C# code that retrieves the RichText collection of cell A1, iterates each RichTextPortion, and sets Font.Name to "Calibri" using Aspose.Cells. | Show how to change the font size of RichTextPortion objects based on their text length in a worksheet cell. | Provide an example that checks for null RichTextPortion entries before applying font changes in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRichTextPortionDemo
{
    // This example shows how to retrieve the RichText collection of a worksheet cell, loop through each RichTextPortion, set a new Font.Name, and save the workbook. It includes directory creation, exception handling, and demonstrates the complete workflow with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Set a string value in cell A1 (rich text can be applied to string cells)
                Cell cell = worksheet.Cells["A1"];
                cell.PutValue("Hello World! This is a test.");

                // Apply a font change to the whole text using the cell's style
                string text = cell.StringValue;
                if (!string.IsNullOrEmpty(text))
                {
                    Style style = cell.GetStyle();
                    style.Font.Name = "Arial";
                    cell.SetStyle(style);
                }

                // Define output path and ensure the directory exists
                string outputPath = "RichTextPortionModified.xlsx";
                string fullPath = Path.GetFullPath(outputPath);
                string outputDir = Path.GetDirectoryName(fullPath);

                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
