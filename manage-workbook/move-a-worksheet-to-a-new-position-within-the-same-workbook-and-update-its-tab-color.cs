// Title: C# – Move Worksheet and Set Tab Color with Aspose.Cells
// Description: Creates a workbook, adds several sheets, moves a specific worksheet to a chosen index using the MoveTo method, applies a green TabColor, and saves the file as MovedSheetWithTabColor.xlsx.
// Keywords: Aspose.Cells MoveTo | Aspose.Cells TabColor | C# reorder worksheet | change worksheet tab color | move sheet index | Aspose.Cells workbook manipulation | set worksheet tab color programmatically | Aspose.Cells example .NET | C# Excel sheet ordering | Aspose.Cells MoveTo method
// Common Searches: Aspose.Cells move worksheet to specific index | How to set worksheet tab color in C# using Aspose.Cells | Reorder Excel sheets with Aspose.Cells .NET | Change Excel tab color programmatically Aspose | Move and color worksheet tabs Aspose.Cells example
// Developer Intent: Reorder a sheet within a workbook and apply a custom tab color using Aspose.Cells for .NET.
// Use Cases: Highlight a summary sheet by moving it to the second tab and coloring it green for quick navigation. | Automate report generation where each section sheet is positioned and colored to reflect its status. | Prepare a template workbook that programmatically arranges and styles tabs before distribution.
// AI Prompts: Generate C# code that moves a worksheet named 'Data' to the third position and sets its tab color to red using Aspose.Cells. | Write a reusable method that accepts a worksheet name, target index, and System.Drawing.Color, then moves the sheet and updates its TabColor. | Explain the interaction between the MoveTo method and the TabColor property for managing worksheet order and appearance in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds several sheets, moves a specific worksheet to a chosen index using the MoveTo method, applies a green TabColor, and saves the file as MovedSheetWithTabColor.xlsx.
    public class MoveWorksheetAndSetTabColor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add some worksheets to have multiple sheets
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Add a new worksheet that we will move
                Worksheet movedSheet = workbook.Worksheets.Add("MovedSheet");

                // Move the worksheet to the desired position (index 1, i.e., second tab)
                movedSheet.MoveTo(1);

                // Set the tab color of the moved worksheet (e.g., green)
                movedSheet.TabColor = Color.Green;

                // Save the workbook to a file
                workbook.Save("MovedSheetWithTabColor.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
