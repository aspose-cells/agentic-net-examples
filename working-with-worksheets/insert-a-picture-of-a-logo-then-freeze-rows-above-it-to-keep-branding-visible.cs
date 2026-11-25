// Title: Insert a PNG logo into cell A1 and freeze the top row using Aspose.Cells for .NET
// Common Searches: how to add a logo image to the first row of an Excel sheet with Aspose.Cells C# | Aspose.Cells C# freeze top row after inserting picture | insert picture at A1 and freeze panes in .NET Excel workbook | set picture size and freeze rows using Aspose.Cells for .NET | load template workbook, add branding image, freeze header row Aspose.Cells
// Tags: add picture to worksheet Aspose.Cells | freeze panes after inserting image Aspose.Cells | set picture dimensions C# Aspose.Cells | insert logo into Excel file .NET | freeze top row for branding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads a workbook (or creates a new one), inserts a PNG logo at cell A1 on the first worksheet, optionally sets its width and height, freezes the top row containing the logo to keep branding visible, and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Define paths for input workbook, logo image and output workbook.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string inputFile = Path.Combine(dataDir, "Template.xlsx");
        string logoFile = Path.Combine(dataDir, "logo.png");
        string outputFile = Path.Combine(dataDir, "Result.xlsx");

        // Verify that the data directory exists.
        if (!Directory.Exists(dataDir))
        {
            Console.WriteLine($"Error: Data directory does not exist: {dataDir}");
            return;
        }

        // Verify that the logo image exists.
        if (!File.Exists(logoFile))
        {
            Console.WriteLine($"Error: Logo image not found at: {logoFile}");
            return;
        }

        Workbook workbook = null;

        // Load an existing workbook if present; otherwise create a new one.
        if (File.Exists(inputFile))
        {
            try
            {
                workbook = new Workbook(inputFile);
                Console.WriteLine($"Loaded workbook from \"{inputFile}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
        }
        else
        {
            workbook = new Workbook();
            Console.WriteLine("No template found – created a new workbook.");
        }

        // Work with the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Branding";

        // Insert the logo picture at cell A1 (zero‑based row/column indices).
        int pictureRow = 0;      // Row index for the picture (A1)
        int pictureColumn = 0;   // Column index for the picture (A1)

        try
        {
            // Add returns the picture index.
            int pictureIndex = sheet.Pictures.Add(pictureRow, pictureColumn, logoFile);
            Picture logoPicture = sheet.Pictures[pictureIndex];

            // Optional: set explicit size for the picture.
            logoPicture.Width = 200;   // pixels
            logoPicture.Height = 80;   // pixels

            Console.WriteLine("Logo picture inserted successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error inserting picture: {ex.Message}");
            return;
        }

        // Freeze the rows that contain the picture (the first row in this case).
        int rowsToFreeze = pictureRow + 1; // Freeze the top row(s) that hold the logo.

        try
        {
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // row/column are zero‑based indices where the split occurs.
            sheet.FreezePanes(rowsToFreeze, 0, rowsToFreeze, 0);
            Console.WriteLine($"Frozen top {rowsToFreeze} row(s) to keep branding visible.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error freezing panes: {ex.Message}");
            return;
        }

        // Ensure the output directory exists.
        try
        {
            string outDir = Path.GetDirectoryName(outputFile);
            if (!Directory.Exists(outDir))
            {
                Directory.CreateDirectory(outDir);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error preparing output directory: {ex.Message}");
            return;
        }

        // Save the modified workbook.
        try
        {
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
