// Title: How to retrieve and print all slicer names from an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, accesses the worksheet's SlicerCollection, and outputs each slicer's Name to the console. | Show an example of iterating over Slicer objects in a worksheet and logging their names using Aspose.Cells for .NET. | Demonstrate how to enumerate slicers in a workbook and display their names without modifying the file.
// Common Searches: aspnet c# enumerate slicers in an Excel workbook using Aspose.Cells | how to get slicer names from a worksheet with Aspose.Cells library | list all slicer objects in an Excel sheet programmatically in C# | Aspose.Cells retrieve slicer collection and print names example
// Tags: Aspose.Cells SlicerCollection enumeration C# | worksheet slicer name extraction Aspose.Cells | log slicer names Aspose.Cells .NET | iterate Excel slicers using Aspose.Cells | C# read slicer objects from workbook

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // The example loads an Excel workbook, accesses the first worksheet's SlicerCollection via Aspose.Cells, loops through each Slicer object, writes its Name to the console, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Get the first worksheet (or any specific worksheet by index/name)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the slicer collection from the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Iterate through each slicer and log its name
            foreach (Slicer slicer in slicers)
            {
                Console.WriteLine("Slicer Name: " + slicer.Name);
            }

            // Save the workbook (optional – can be the same or a new file)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
