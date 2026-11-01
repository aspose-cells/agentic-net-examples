// Title: Adjust column width, row height, and reposition a shape then save the workbook as XLSX while preserving existing data with Aspose.Cells for .NET
// AI Prompts: Load an existing workbook (or create a new one if the file is missing), set column A width to 20 units and row 1 height to 30 units, move the first shape 5 points right and 5 points down, and save the result as a new XLSX file using Aspose.Cells. | Apply geometry changes to the first worksheet—column width, row height, and shape position—without altering any cell values, then export the workbook in XLSX format. | Implement fallback logic that creates a default workbook when the input file does not exist, performs column/row sizing and shape relocation, and persists the changes to a separate output file.
// Common Searches: how to change column width and row height with Aspose.Cells C# without affecting data | move a shape programmatically in an Excel worksheet using Aspose.Cells .NET | save workbook after modifying shape geometry while keeping original content Aspose.Cells | create a new workbook when input file is missing Aspose.Cells C# example | adjust Excel shape position by points using Aspose.Cells API
// Tags: set column width Aspose.Cells C# | set row height Aspose.Cells C# | reposition shape Aspose.Cells | save workbook as XLSX Aspose.Cells | fallback workbook creation Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The code loads an existing XLSX file (or creates a default workbook if the file is absent), adjusts column A width to 20, row 1 height to 30, moves the first shape 5 points right and down, and saves the modified workbook to a new XLSX file while preserving all original worksheet data.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; if not, create a new workbook as a fallback.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Geometry changes -----
            // Adjust column width and row height.
            sheet.Cells.SetColumnWidth(0, 20); // Column A width.
            sheet.Cells.SetRowHeight(0, 30);   // Row 1 height.

            // Reposition the first shape on the sheet, if any.
            if (sheet.Shapes.Count > 0)
            {
                Shape shape = sheet.Shapes[0];
                // Move shape 5 points right and 5 points down (1 point = 9525 EMU).
                shape.Left = 5 * 9525;
                shape.Top = 5 * 9525;
            }

            // Save the workbook to a new XLSX file.
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
