// Title: How to apply a diagonal crosshatch fill pattern to placeholder cells in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set diagonal crosshatch pattern for cells containing placeholder markers | C# code to style Excel cells with {{}} placeholders using BackgroundType.DiagonalCrosshatch | Apply custom fill pattern to specific cells in a .NET Excel file with Aspose.Cells | Identify placeholder text cells and change their background pattern in Aspose.Cells | Set foreground color for diagonal crosshatch fill in an Aspose.Cells workbook
// Tags: diagonal crosshatch fill Aspose.Cells | placeholder cell styling .NET | BackgroundType.DiagonalCrosshatch usage | Excel cell pattern formatting C# | custom fill pattern Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, scans all used cells in the first worksheet, detects cells whose text contains "{{" placeholders, and applies a diagonal crosshatch fill pattern with a light‑gray foreground color by modifying the cell's Style before saving the file.
class ApplyDiagonalCrosshatch
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                // Load without password (assumes file is not password‑protected)
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a sample placeholder cell.
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("{{SamplePlaceholder}}");
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in worksheet.Cells)
            {
                // Identify placeholder text cells (e.g., containing "{{" and "}}")
                if (cell.Type == CellValueType.IsString && cell.StringValue.Contains("{{"))
                {
                    // Retrieve the current style of the cell
                    Style style = cell.GetStyle();

                    // Set the fill pattern to diagonal crosshatch
                    style.Pattern = BackgroundType.DiagonalCrosshatch;

                    // Optionally set a foreground color for the pattern
                    style.ForegroundColor = Color.LightGray;

                    // Apply the modified style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
