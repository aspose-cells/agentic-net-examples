// Title: Programmatically assign the CalculateTotals VBA macro to a button form control in an Aspose.Cells .NET workbook
// Common Searches: how to link a VBA macro to a button shape using Aspose.Cells for .NET | assign macro to button control in a macro-enabled workbook with C# Aspose.Cells | add a form button that runs CalculateTotals macro in an Aspose.Cells Xlsm file | set button.Macro property Aspose.Cells C# example | create button form control with macro reference in an Aspose.Cells workbook
// Tags: assign VBA macro to button shape Aspose.Cells | create form button with macro reference C# | macro-enabled Xlsm workbook button linking Aspose.Cells | AddButton API usage for macro-enabled files | Aspose.Cells button control VBA integration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The code creates a new workbook, adds a button form control at a specified cell location, sets its caption to "Calculate Totals", shows how to assign the CalculateTotals VBA macro via the button's Macro property (not available in older Aspose.Cells versions), and saves the workbook as a macro‑enabled .xlsm file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define button position and size (row, column, top offset, left offset, height, width)
            int row = 2;          // zero‑based row index
            int column = 1;       // zero‑based column index
            int top = 5;          // pixels from the top of the cell
            int left = 5;         // pixels from the left of the cell
            int height = 30;      // button height in pixels
            int width = 100;      // button width in pixels

            // Add a button form control to the worksheet
            Button button = worksheet.Shapes.AddButton(row, column, top, left, height, width);
            button.Text = "Calculate Totals";

            // NOTE: The Macro property is not available in the current Aspose.Cells version.
            // If macro support is required, ensure you are using a version that provides the Macro property
            // and uncomment the following line:
            // button.Macro = "CalculateTotals";

            // Define output file path
            string outputPath = "WorkbookWithButton.xlsm";

            // Save the workbook as a macro‑enabled file
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
