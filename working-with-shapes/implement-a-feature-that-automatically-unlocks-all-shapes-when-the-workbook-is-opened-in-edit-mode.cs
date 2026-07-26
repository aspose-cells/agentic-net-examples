// Title: C# – Unlock All Shapes on Workbook Open with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through every worksheet and shape, sets each shape's IsLocked property to false, and saves the file so all shapes are editable when the workbook is opened in edit mode.
// Keywords: Aspose.Cells | C# | unlock shapes | IsLocked false | Excel shape protection | edit mode | batch shape unlock | programmatic shape unlocking | Aspose.Cells for .NET
// Common Searches: Aspose.Cells unlock all shapes C# | set shape IsLocked false Aspose.Cells | remove shape protection programmatically Excel | unlock shapes when opening workbook in edit mode | C# code to unlock Excel shapes with Aspose
// Developer Intent: Programmatically remove shape locks so users can edit any shape immediately after opening the workbook.
// Use Cases: Prepare a template where end‑users must move or edit shapes without manual unlocking. | Batch‑process multiple workbooks to clear shape protection before distribution. | Provide a macro‑free solution that ensures all shapes are unlocked automatically on workbook open.
// AI Prompts: Write C# code using Aspose.Cells that unlocks every shape in a workbook when it is opened. | Show how to unlock shapes only on selected worksheets while keeping others locked, using the IsLocked property. | Create an example that adds an event‑handler‑like routine to automatically unlock shapes on workbook load with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, iterates through every worksheet and shape, sets each shape's IsLocked property to false, and saves the file so all shapes are editable when the workbook is opened in edit mode.
    public class UnlockAllShapesOnOpen
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple workbook if missing
                if (!File.Exists(inputPath))
                {
                    var tempWb = new Workbook();
                    var tempSheet = tempWb.Worksheets[0];
                    // Add a sample shape so there is something to unlock
                    tempSheet.Shapes.AddShape(MsoDrawingType.Rectangle, 1, 1, 0, 0, 100, 50);
                    tempWb.Save(inputPath);
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Unlock all shapes in every worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        shape.IsLocked = false;
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

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            UnlockAllShapesOnOpen.Run();
        }
    }
}
