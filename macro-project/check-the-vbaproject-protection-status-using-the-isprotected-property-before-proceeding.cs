// Title: How to check VBA project protection status with Aspose.Cells for .NET before modifying a workbook
// Common Searches: C# Aspose.Cells detect if a VBA project is password protected | skip adding worksheets when workbook VbaProject.IsProtected is true | verify VBA macro protection in an Excel file using Aspose.Cells .NET | prevent workbook changes if the embedded VBA project is locked with Aspose.Cells
// Tags: Aspose.Cells VBA project protection check | C# detect protected VBA macro | skip workbook modification when VBA locked | Aspose.Cells VbaProject.IsProtected usage | prevent processing of Excel file with protected VBA

using Aspose.Cells;
using System;

// The example loads an Excel workbook, checks whether a VBA project exists and whether it is protected via the workbook.VbaProject.IsProtected property, aborts further processing if protection is found, and otherwise adds a new worksheet and saves the file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Verify that a VBA project exists and check its protection status
        if (workbook.VbaProject != null && workbook.VbaProject.IsProtected)
        {
            // If the VBA project is protected, stop further processing
            Console.WriteLine("The VBA project is protected. Operation aborted.");
            return;
        }

        // Continue with desired operations (example: add a new worksheet)
        workbook.Worksheets.Add("NewSheet");

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
