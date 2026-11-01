// Title: Unlock the 'Signature' shape and set it to free‑floating placement in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an XLSX workbook, locates the shape named "Signature" on the first worksheet, disables its lock, changes its placement to free‑floating, and saves the result with Aspose.Cells. | Generate a program using Aspose.Cells that searches for a named shape, sets IsLocked = false, sets Placement = FreeFloating, and writes the updated file in C#. | Create a C# example that unlocks the "Signature" shape, allows it to be resized and moved freely, and persists the changes to a new Excel file.
// Common Searches: Aspose.Cells C# unlock shape named Signature and change placement | set shape to free floating placement using Aspose.Cells .NET | how to modify IsLocked property of a shape in an Excel workbook with Aspose.Cells | C# code to find and edit a specific shape in an XLSX file using Aspose.Cells | unlock and reposition Excel shape programmatically with Aspose.Cells
// Tags: disable shape lock Aspose.Cells C# | apply freefloating placement Aspose.Cells | locate shape by name Aspose.Cells worksheet | edit shape properties Excel Aspose.Cells | allow shape resizing after unlocking Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads 'input.xlsx', searches the first worksheet for a shape called "Signature", sets its IsLocked property to false, changes its Placement to FreeFloating, and saves the modified workbook as 'output.xlsx', with error handling for missing files and exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Locate the shape named "Signature"
                int signatureIndex = -1;
                for (int i = 0; i < worksheet.Shapes.Count; i++)
                {
                    if (worksheet.Shapes[i].Name == "Signature")
                    {
                        signatureIndex = i;
                        break;
                    }
                }

                // If the shape exists, unlock it and allow free movement
                if (signatureIndex != -1)
                {
                    Shape signatureShape = worksheet.Shapes[signatureIndex];

                    // Unlock the shape so it can be edited
                    signatureShape.IsLocked = false;

                    // Allow the shape to be moved freely
                    signatureShape.Placement = PlacementType.FreeFloating;

                    // Note: Aspose.Cells Shape does not expose a LockAspectRatio property.
                    // The shape can be resized after unlocking; no additional code is required.
                }
                else
                {
                    Console.WriteLine("Warning: Shape named \"Signature\" was not found.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
