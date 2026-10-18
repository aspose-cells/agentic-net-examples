// Title: How to apply a diagonal cross‑hatch fill pattern to a header row (A1:D1) in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set diagonal crosshatch background for a range of cells | C# apply custom fill pattern to header row in Excel workbook with Aspose.Cells | How to change cell style to BackgroundType.DiagonalCrosshatch using Aspose.Cells | Set foreground and background colors for patterned cell style in Aspose.Cells .NET | Create header cells with dense hatch fill in Excel using Aspose.Cells API
// Tags: apply diagonal crosshatch pattern Aspose.Cells | header row cell style Aspose.Cells C# | BackgroundType.DiagonalCrosshatch usage | set cell foreground background colors Aspose.Cells | style range cells Aspose.Cells .NET

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates creating a workbook, selecting the A1:D1 range, and applying a diagonal cross‑hatch fill pattern with LightGray foreground and White background to each header cell using Aspose.Cells for .NET, then saving the file as HeaderPattern.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the header range (e.g., cells A1 to D1)
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("A1:D1");

            // Apply a dense hatch-like fill pattern to each header cell
            foreach (Cell cell in headerRange)
            {
                // Retrieve the current style of the cell
                Style style = cell.GetStyle();

                // Set the background pattern to a crosshatch style
                style.Pattern = BackgroundType.DiagonalCrosshatch;

                // Set foreground and background colors
                style.ForegroundColor = Color.LightGray;
                style.BackgroundColor = Color.White;

                // Apply the modified style back to the cell
                cell.SetStyle(style);
            }

            // Save the workbook to a file
            workbook.Save("HeaderPattern.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
