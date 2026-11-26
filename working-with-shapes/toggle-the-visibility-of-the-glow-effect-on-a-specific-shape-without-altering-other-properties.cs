// Title: How to toggle a shape’s glow effect on and off in Aspose.Cells for .NET without changing other properties
// Common Searches: Aspose.Cells C# hide shape glow by setting size to zero | how to toggle glow effect of a rectangle shape in an Excel workbook using Aspose.Cells | restore previous glow size after hiding it with Aspose.Cells .NET | change only the glow effect of a shape while keeping other formatting unchanged in Aspose.Cells | toggle shape glow visibility on and off without affecting other shape properties in C#
// Tags: shape glow size control Aspose.Cells | hide shape glow Aspose.Cells .NET | restore shape glow properties Aspose.Cells | preserve shape formatting Aspose.Cells | C# Aspose.Cells glow effect handling

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace ToggleGlowEffectDemo
{
    // The example creates a workbook, adds a rectangle shape, applies a purple glow, saves the file, hides the glow by setting its Size to 0 while preserving other properties, saves again, restores the original glow size, and saves a final workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path where the output files will be saved
            string outputFolder = AppDomain.CurrentDomain.BaseDirectory;

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left offsetX, upper left offsetY, width, height
            Shape shape = sheet.Shapes.AddRectangle(1, 0, 0, 0, 200, 100);
            shape.Name = "DemoRectangle";

            // Configure initial glow effect (visible)
            GlowEffect glow = shape.Glow;
            CellsColor glowColor = workbook.CreateCellsColor();
            glowColor.Color = Color.Purple; // Set glow color
            glow.Color = glowColor;
            glow.Size = 12;               // Radius in points (non‑zero makes it visible)
            glow.Transparency = 0.3;      // 30% transparent

            // Save workbook with glow visible
            string visiblePath = System.IO.Path.Combine(outputFolder, "GlowVisible.xlsx");
            workbook.Save(visiblePath);
            Console.WriteLine($"Saved workbook with glow visible: {visiblePath}");

            // Toggle glow visibility off by setting Size to 0 (other properties stay unchanged)
            double previousSize = glow.Size; // Remember current size to restore later
            glow.Size = 0;                   // Size 0 effectively hides the glow

            // Save workbook with glow hidden
            string hiddenPath = System.IO.Path.Combine(outputFolder, "GlowHidden.xlsx");
            workbook.Save(hiddenPath);
            Console.WriteLine($"Saved workbook with glow hidden: {hiddenPath}");

            // Toggle glow visibility back on by restoring the previous size
            glow.Size = previousSize; // Restore original size

            // Save workbook with glow visible again
            string restoredPath = System.IO.Path.Combine(outputFolder, "GlowRestored.xlsx");
            workbook.Save(restoredPath);
            Console.WriteLine($"Saved workbook with glow restored: {restoredPath}");

            // Clean up
            workbook.Dispose();

            Console.WriteLine("Glow toggle demonstration completed.");
        }
    }
}
