// Title: How to retrieve and log the address of a dynamic named range (SalesData) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get address of a named range from workbook | retrieve RefersTo value of dynamic named range in Aspose.Cells | C# code to read named range address in Excel with Aspose.Cells | log named range address to console using Aspose.Cells .NET | check if named range exists and obtain its reference in Aspose.Cells
// Tags: Aspose.Cells named range RefersTo extraction | C# workbook named range address retrieval | dynamic range address logging Aspose.Cells | Workbook.Worksheets.Names API usage | Excel named range existence check Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, accesses the dynamic named range "SalesData" via Workbook.Worksheets.Names, verifies its presence, extracts its RefersTo address (e.g., Sheet1!$A$1:$D$10), and writes the address to the console.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the named range "SalesData"
        Name salesRange = workbook.Worksheets.Names["SalesData"];

        if (salesRange != null)
        {
            // Get the address the name refers to (e.g., Sheet1!$A$1:$D$10)
            string address = salesRange.RefersTo;

            // Log the address
            Console.WriteLine($"SalesData range address: {address}");
        }
        else
        {
            Console.WriteLine("Named range 'SalesData' not found.");
        }
    }
}
