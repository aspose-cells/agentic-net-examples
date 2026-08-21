// Title: C# – Log FitToPagesWide changes in Aspose.Cells worksheet PageSetup
// Description: Demonstrates how to read the default FitToPagesWide setting of a worksheet, output the original value, modify it using the FitToPagesWide property and the SetFitToPages method, log each new value, and save the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | FitToPagesWide | PageSetup | SetFitToPages | log page scaling | worksheet print settings | track changes
// Common Searches: Aspose.Cells log FitToPagesWide value | C# get default FitToPagesWide | change FitToPagesWide and capture previous value | SetFitToPages example Aspose.Cells | how to audit worksheet page setup in .NET
// Developer Intent: Capture the initial FitToPagesWide setting and record each modification made to the worksheet's print scaling.
// Use Cases: Audit printing configuration before and after scaling adjustments. | Generate a change history for dynamic report layouts that alter page width. | Validate batch updates of page setup across multiple worksheets.
// AI Prompts: Create a C# routine that writes the original and updated FitToPagesWide values to a log file with timestamps. | Show how to wrap FitToPagesWide changes in an event handler that records every modification for a worksheet. | Explain the difference between setting FitToPagesWide directly and using SetFitToPages, including how to log both outcomes.

using System;
using Aspose.Cells;

namespace AsposeCellsFitToPagesWideLogger
{
    // Demonstrates how to read the default FitToPagesWide setting of a worksheet, output the original value, modify it using the FitToPagesWide property and the SetFitToPages method, log each new value, and save the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Log the original FitToPagesWide value (default is 1)
            int originalFitToPagesWide = pageSetup.FitToPagesWide;
            Console.WriteLine($"Original FitToPagesWide: {originalFitToPagesWide}");

            // Change the FitToPagesWide value using the property
            pageSetup.FitToPagesWide = 2;

            // Log the updated FitToPagesWide value
            int updatedFitToPagesWide = pageSetup.FitToPagesWide;
            Console.WriteLine($"Updated FitToPagesWide (property): {updatedFitToPagesWide}");

            // Alternatively, change the value using SetFitToPages method
            pageSetup.SetFitToPages(3, pageSetup.FitToPagesTall);
            Console.WriteLine($"Updated FitToPagesWide (SetFitToPages): {pageSetup.FitToPagesWide}");

            // Save the workbook (save rule)
            workbook.Save("FitToPagesWideLogDemo.xlsx");
        }
    }
}
