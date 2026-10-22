// Title: Insert a PNG logo at the top of an Excel worksheet and freeze the rows above it with Aspose.Cells for .NET (C#)
// AI Prompts: Place a PNG logo in cell A1 of the first worksheet using Aspose.Cells for .NET and lock the rows that contain the logo so the branding remains visible while scrolling. | After inserting a picture, obtain its UpperLeftRow property and call the appropriate method to freeze the rows above the image in a C# Aspose.Cells workbook. | Rewrite the sample to read the logo from a MemoryStream instead of a file path while still applying the row‑freezing logic with Aspose.Cells.
// Common Searches: Aspose.Cells C# insert image at top of sheet and keep header frozen | freeze rows after adding picture with Aspose.Cells .NET | how to programmatically add a logo to Excel and freeze rows using Aspose.Cells | C# Aspose.Cells freeze rows based on picture position
// Tags: insert image into Excel worksheet Aspose.Cells C# | apply FreezePanes API to lock rows Aspose.Cells | logo branding in Excel workbook using Aspose.Cells | load picture from MemoryStream Aspose.Cells | determine picture UpperLeftRow property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, inserts a PNG logo at the top‑left cell, calculates the picture's starting row, freezes all rows above that row using FreezePanes, and saves the file as OutputWithLogo.xlsx.
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

            // Path to the logo image
            string logoPath = "logo.png";

            // Verify that the logo file exists
            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Logo file not found: {logoPath}");
                return;
            }

            // Insert the logo picture at the top‑left corner (row 0, column 0)
            int pictureIndex = sheet.Pictures.Add(0, 0, logoPath);

            // Retrieve the inserted picture to determine its position
            Picture logo = sheet.Pictures[pictureIndex];

            // Freeze rows above the picture (including the row where the picture starts)
            int freezeRow = logo.UpperLeftRow + 1; // UpperLeftRow is the starting row of the picture

            // Freeze panes: row, column, number of rows to freeze, number of columns to freeze
            sheet.FreezePanes(freezeRow, 0, freezeRow, 0);

            // Save the workbook to a file
            string outputPath = "OutputWithLogo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
