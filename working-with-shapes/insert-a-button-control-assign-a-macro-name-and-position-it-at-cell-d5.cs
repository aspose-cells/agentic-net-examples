// Title: Add a Button shape to cell D5, assign a VBA macro, and save the workbook using Aspose.Cells for .NET
// Common Searches: asp.net add button shape to cell D5 with macro using Aspose.Cells | c# Aspose.Cells place button at D5 and set MacroName property | how to assign a VBA macro to a button control in Aspose.Cells worksheet | Aspose.Cells example for creating an Excel button that runs MyMacro() | save an Excel file with a button that triggers a macro using Aspose.Cells C#
// Tags: add button shape Aspose.Cells C# | set button MacroName Aspose.Cells | position button at cell D5 Aspose.Cells | button control Excel VBA macro Aspose.Cells | save workbook with button macro Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsButtonMacroDemo
{
    // The sample creates a new workbook, adds a Button shape positioned at cell D5 (row 4, column 3) with defined height and width, assigns the MacroName "MyMacro()" and a caption, then saves the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the target cell for the button (D5)
                // Row and column indices are zero‑based: D -> column 3, 5 -> row 4
                int targetRow = 4;      // Row 5 (zero‑based)
                int targetColumn = 3;   // Column D (zero‑based)

                // Define button size (in pixels)
                int buttonHeight = 30;
                int buttonWidth = 80;

                // Offsets within the cell (top and left) – set to 0 for top‑left alignment
                int offsetTop = 0;
                int offsetLeft = 0;

                // Add a Button control to the worksheet
                // Signature: AddButton(int upperLeftRow, int top, int upperLeftColumn, int left, int height, int width)
                Shape buttonShape = worksheet.Shapes.AddButton(
                    targetRow,          // upper left row index
                    offsetTop,          // vertical offset in pixels
                    targetColumn,       // upper left column index
                    offsetLeft,         // horizontal offset in pixels
                    buttonHeight,       // height in pixels
                    buttonWidth);       // width in pixels

                // Set the macro name that will be called when the button is clicked
                // The macro name must match a VBA macro defined in the workbook (e.g., "MyMacro")
                buttonShape.MacroName = "MyMacro()";

                // Optionally set the button caption
                buttonShape.Text = "Run Macro";

                // Save the workbook to a file
                string outputPath = "ButtonWithMacro.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("Button added at cell D5 with macro name 'MyMacro()'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
