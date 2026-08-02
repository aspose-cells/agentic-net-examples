// Title: Decrease a Shape’s Z‑Order by Three in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add a rectangle shape, read its ZOrderPosition, subtract three (with a floor of zero), assign the new value, verify the change, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# shape Z-order | set ZOrderPosition | move shape backward | adjust shape stacking order | Aspose.Cells shape order example | Excel shape Z-order .NET | programmatic shape layering | C# Aspose.Cells ZOrderPosition | shape order manipulation | Excel drawing order
// Common Searches: how to lower a shape's Z-order in Aspose.Cells | C# code to move an Excel shape backward by three positions | Aspose.Cells set shape ZOrderPosition to zero | verify shape Z-order after change in .NET | adjust shape stacking order programmatically in Excel
// Developer Intent: Programmatically lower a shape’s Z-order by three slots while preventing negative indices.
// Use Cases: Place a newly added rectangle behind existing charts or images. | Prepare a worksheet for export where certain shapes must stay in the background. | Implement dynamic layering based on user preferences or data-driven rules.
// AI Prompts: Generate C# code with Aspose.Cells that moves a specific shape three positions backward in Z-order and clamps the result at zero. | Create an example that iterates over all worksheet shapes and reduces each ZOrderPosition by a given offset without dropping below zero. | Explain how to read, modify, and validate a shape's ZOrderPosition after adjustment using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsZOrderAdjustment
{
    // Shows how to create a workbook, add a rectangle shape, read its ZOrderPosition, subtract three (with a floor of zero), assign the new value, verify the change, and save the file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                Shape shape = worksheet.Shapes.AddRectangle(10, 10, 100, 100, 0, 0);

                // Retrieve the current Z-order position of the shape
                int currentZOrder = shape.ZOrderPosition;
                Console.WriteLine("Current Z-order position: " + currentZOrder);

                // Calculate a new Z-order position, ensuring it stays within valid bounds
                int newZOrder = Math.Max(0, currentZOrder - 3);
                if (newZOrder != currentZOrder)
                {
                    // Apply the new Z-order position
                    shape.ZOrderPosition = newZOrder;
                }

                // Verify the new Z-order position
                int verifiedZOrder = shape.ZOrderPosition;
                Console.WriteLine("New Z-order position after adjustment: " + verifiedZOrder);

                // Save the workbook to verify the shape is retained
                workbook.Save("ZOrderAdjusted.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
