// Title: Create a custom Arial 12 font style and apply it to cells A1:C3 with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set font name and size for a range of cells | Apply a style with only font name and size to cells A1:C3 using Aspose.Cells .NET | Using StyleFlag to change font attributes in Aspose.Cells workbook | Create and apply custom style to a cell area in Aspose.Cells for .NET | Save a styled workbook with Arial 12 font using Aspose.Cells C#
// Tags: custom font style creation Aspose.Cells C# | apply font style to cell range Aspose.Cells | StyleFlag selective font formatting Aspose.Cells | Arial 12 font setting Aspose.Cells .NET | styled workbook export Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, defines a style with the Arial font at size 12, selects the A1:C3 range, and uses a StyleFlag to apply only the font name and size to that range before saving the workbook as StyledWorkbook.xlsx.
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

            // Define the range of cells to which the style will be applied (A1:C3)
            CellArea area = CellArea.CreateCellArea("A1", "C3");

            // Configure which style attributes to apply (font name and size)
            StyleFlag flag = new StyleFlag
            {
                FontName = true,
                FontSize = true
            };

            // Apply the style to the specified range using a Range object
            int rows = area.EndRow - area.StartRow + 1;
            int columns = area.EndColumn - area.StartColumn + 1;
            AsposeRange range = sheet.Cells.CreateRange(area.StartRow, area.StartColumn, rows, columns);
            range.ApplyStyle(style, flag);

            // Determine output path
            string outputPath = "StyledWorkbook.xlsx";

            // Ensure the output directory exists (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
