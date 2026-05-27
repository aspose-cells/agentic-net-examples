using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class RetrievePaperSizeAndCalculateScale
{
    static void Main()
    {
        // Create a new workbook or load an existing one
        Workbook workbook = new Workbook(); // or new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access the PageSetup object of the worksheet
        PageSetup pageSetup = sheet.PageSetup;

        // Retrieve the paper width and height in inches (read‑only properties)
        double paperWidthInches = pageSetup.PaperWidth;
        double paperHeightInches = pageSetup.PaperHeight;

        Console.WriteLine($"Paper Width (inches): {paperWidthInches}");
        Console.WriteLine($"Paper Height (inches): {paperHeightInches}");

        // Example: calculate custom scaling factors based on a target size
        // Suppose we want the printed area to fit within 6 inches width and 8 inches height
        double targetWidthInches = 6.0;
        double targetHeightInches = 8.0;

        // Scaling factors (greater than 1 means enlarge, less than 1 means shrink)
        double scaleX = targetWidthInches / paperWidthInches;
        double scaleY = targetHeightInches / paperHeightInches;

        Console.WriteLine($"Scale X (width factor): {scaleX:F3}");
        Console.WriteLine($"Scale Y (height factor): {scaleY:F3}");

        // Optionally, apply a custom paper size that matches the target dimensions
        // This demonstrates using the CustomPaperSize method
        pageSetup.CustomPaperSize(targetWidthInches, targetHeightInches);
        Console.WriteLine($"Custom paper size set to {targetWidthInches}\" x {targetHeightInches}\"");

        // Save the workbook to verify the changes (if needed)
        workbook.Save("OutputWithCustomPaperSize.xlsx");
    }
}