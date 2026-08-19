// Title: Make all slicers non‑printable in an Excel workbook using Aspose.Cells for .NET
// Description: Loads an existing workbook, walks through each worksheet, accesses the SlicerCollection, sets the IsPrintable property of every Slicer to false, and saves the result as a new file.
// Keywords: Aspose.Cells | C# slicer printing | IsPrintable false | disable slicer print | iterate slicers .NET | Excel slicer non‑printable | Aspose.Cells SlicerCollection | Excel workbook automation
// Common Searches: Aspose.Cells set slicer IsPrintable false | how to hide slicers from printing in C# | disable slicer printing for all worksheets | make Excel slicers non‑printable programmatically | iterate over slicers with Aspose.Cells
// Developer Intent: Turn off printing for every slicer in the workbook.
// Use Cases: Prepare a printable report where slicers should be omitted. | Generate PDFs from Excel files without slicer graphics. | Create a template that automatically suppresses slicer output on print.
// AI Prompts: Provide C# code that iterates through all worksheets and sets each slicer's IsPrintable property to false with Aspose.Cells. | Show an example of disabling slicer printing and saving the workbook as a new file using Aspose.Cells for .NET. | Explain how to check for the presence of slicers before modifying their printable flag in an Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an existing workbook, walks through each worksheet, accesses the SlicerCollection, sets the IsPrintable property of every Slicer to false, and saves the result as a new file.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path to the destination workbook
        string outputPath = "output.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Iterate through each slicer and set its printable flag to false
            foreach (Slicer slicer in slicers)
            {
                slicer.IsPrintable = false;
            }
        }

        // Save the modified workbook (save rule)
        workbook.Save(outputPath);
    }
}
