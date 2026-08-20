// Title: Toggle Shape Texture Tiling in Excel with Aspose.Cells for .NET
// Description: A C# routine that loads or creates an Excel workbook, adds a rectangle shape with a built‑in texture if none exists, ensures the shape uses texture fill, flips the IsTiling flag of the texture fill on each run, writes a status message, and saves the workbook so the tiling state persists.
// Keywords: Aspose.Cells | C# | .NET | Excel shape texture fill | texture tiling | IsTiling property | toggle texture tiling | shape fill type | programmatic Excel graphics | persist shape settings
// Common Searches: how to toggle texture tiling on an Excel shape using Aspose.Cells | C# code to enable or disable texture fill tiling in a worksheet | Aspose.Cells toggle IsTiling property example | persist shape texture tiling state across workbook saves | macro‑like routine for texture tiling in Excel with .NET
// Developer Intent: Flip the IsTiling flag of a shape’s texture fill each time the code runs, creating or updating the workbook as needed.
// Use Cases: Automatically switch a shape’s texture between tiled and non‑tiled on successive executions. | Create a workbook that remembers the last tiling state of a shape and updates it on demand. | Provide a macro‑style feature for end‑users to toggle texture tiling without manual Excel actions.
// AI Prompts: Generate a C# method using Aspose.Cells that loads a workbook, adds a rectangle with a texture if missing, toggles its IsTiling property, and saves the file. | Write error‑handling code for the texture‑tiling toggle routine to manage missing files, absent shapes, or unsupported fill types. | Show how to use a different built‑in texture (e.g., TextureType.Wood) and set an initial tiling state before toggling it.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsToggleTextureTiling
{
    // A C# routine that loads or creates an Excel workbook, adds a rectangle shape with a built‑in texture if none exists, ensures the shape uses texture fill, flips the IsTiling flag of the texture fill on each run, writes a status message, and saves the workbook so the tiling state persists.
    public class ToggleTilingRoutine
    {
        // Path to the workbook that stores the shape and its tiling state
        private const string WorkbookPath = "ToggleTextureTilingDemo.xlsx";

        public static void Run()
        {
            Workbook workbook;

            // Load existing workbook if it exists, otherwise create a new one
            if (File.Exists(WorkbookPath))
            {
                // Load rule
                workbook = new Workbook(WorkbookPath);
            }
            else
            {
                // Create rule
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape
                Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
                // Configure the shape to use texture fill
                shape.Fill.FillType = FillType.Texture;
                // Use a built‑in texture type
                shape.Fill.TextureFill.Type = TextureType.BlueTissuePaper;
                // Initial tiling state (false)
                shape.Fill.TextureFill.IsTiling = false;
            }

            // Access the first worksheet and the first shape
            Worksheet worksheet = workbook.Worksheets[0];
            if (worksheet.Shapes.Count == 0)
            {
                Console.WriteLine("No shapes found in the worksheet.");
                return;
            }

            Shape targetShape = worksheet.Shapes[0];

            // Ensure the shape uses texture fill
            targetShape.Fill.FillType = FillType.Texture;

            // Toggle the IsTiling property
            bool currentTiling = targetShape.Fill.TextureFill.IsTiling;
            targetShape.Fill.TextureFill.IsTiling = !currentTiling;

            Console.WriteLine($"Texture tiling was {(currentTiling ? "enabled" : "disabled")}, now it is {(!currentTiling ? "enabled" : "disabled")}.");

            // Save rule
            workbook.Save(WorkbookPath);
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            ToggleTilingRoutine.Run();
        }
    }
}
