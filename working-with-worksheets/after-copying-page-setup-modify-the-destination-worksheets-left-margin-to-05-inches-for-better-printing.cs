// Title: Copy page setup and set left margin to 0.5 inches in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to copy the PageSetup from one worksheet to another using Aspose.Cells, then adjust the destination sheet's left margin to 0.5 inches before saving the workbook.
// Keywords: Aspose.Cells | C# | .NET | copy page setup | worksheet left margin | margin adjustment | CopyOptions | printing settings | PageSetup.Copy | Excel automation
// Common Searches: Aspose.Cells copy page setup C# | set worksheet left margin 0.5 inches Aspose.Cells | adjust page margins after copying in .NET | how to change left margin after PageSetup.Copy | copy page layout between sheets Aspose.Cells
// Developer Intent: Copy a worksheet's PageSetup to another sheet and then set the destination sheet's left margin to 0.5 inches.
// Use Cases: Apply a standard page layout to multiple sheets while customizing the left margin for binding or narrow printing. | Create report templates where the first sheet uses default margins and subsequent sheets require a reduced left margin. | Automate workbook generation that reuses page settings across sheets but needs a specific left margin for each sheet.
// AI Prompts: Show C# code that copies a worksheet's PageSetup to another worksheet using Aspose.Cells and then sets the left margin to 0.5 inches. | Explain how CopyOptions work with PageSetup.Copy and how to modify margins after the copy operation. | Provide a step‑by‑step guide to adjust only the left margin of a destination worksheet after copying page setup in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to copy the PageSetup from one worksheet to another using Aspose.Cells, then adjust the destination sheet's left margin to 0.5 inches before saving the workbook.
    public class CopyPageSetupAndAdjustMargin
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a second worksheet (the destination)
                workbook.Worksheets.Add();

                // Access source and destination worksheets
                Worksheet sourceSheet = workbook.Worksheets[0];
                Worksheet destSheet = workbook.Worksheets[1];

                // Configure some page setup settings on the source worksheet
                PageSetup sourceSetup = sourceSheet.PageSetup;
                sourceSetup.PaperSize = PaperSizeType.PaperA4;
                sourceSetup.Orientation = PageOrientationType.Portrait;
                sourceSetup.LeftMarginInch = 1.0;   // initial left margin (in inches)
                sourceSetup.RightMarginInch = 0.75;
                sourceSetup.TopMarginInch = 0.5;
                sourceSetup.BottomMarginInch = 0.5;

                // Copy the page setup from source to destination using default copy options
                destSheet.PageSetup.Copy(sourceSetup, new CopyOptions());

                // After copying, adjust the left margin of the destination worksheet to 0.5 inches
                destSheet.PageSetup.LeftMarginInch = 0.5;

                // Save the workbook to demonstrate the result
                workbook.Save("CopyPageSetupAdjustedMargin.xlsx");
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
            CopyPageSetupAndAdjustMargin.Run();
        }
    }
}
