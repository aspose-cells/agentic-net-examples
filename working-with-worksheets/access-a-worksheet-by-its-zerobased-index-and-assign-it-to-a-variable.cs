// Title: Aspose.Cells C# – Access Worksheet by Zero‑Based Index and Assign to Variable
// Description: Demonstrates how to create a Workbook, retrieve the first Worksheet using the zero‑based index (workbook.Worksheets[0]), assign it to a Worksheet variable, write a value to cell A1, and save the file as WorksheetByIndex.xlsx.
// Keywords: Aspose.Cells worksheet index | C# get worksheet by index | Aspose.Cells Worksheets[0] | assign worksheet variable | write cell A1 Aspose.Cells | .NET Aspose.Cells example
// Common Searches: Aspose.Cells get first worksheet | How to access worksheet by index C# | Aspose.Cells Worksheets collection example | Assign worksheet to variable Aspose.Cells | Write value to cell after retrieving worksheet
// Developer Intent: Retrieve a worksheet from a workbook using its zero‑based index and store it in a variable for further manipulation.
// Use Cases: Read data from a specific sheet after selecting it by index. | Update cell values, formulas, or formatting on a sheet obtained via index. | Iterate through all worksheets by index to apply consistent styles or headers. | Export or copy a particular worksheet identified by its position in the workbook.
// AI Prompts: Generate C# code that opens an existing workbook and accesses the second worksheet by index using Aspose.Cells. | Show how to loop through every worksheet in a workbook by index and set a header row in each sheet with Aspose.Cells. | Provide an example that copies data from a worksheet accessed by index to another worksheet within the same workbook using Aspose.Cells.

using System;
using Aspose.Cells;

namespace WorksheetAccessExample
{
    // Demonstrates how to create a Workbook, retrieve the first Worksheet using the zero‑based index (workbook.Worksheets[0]), assign it to a Worksheet variable, write a value to cell A1, and save the file as WorksheetByIndex.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet by its zero‑based index and assign it to a variable
            Worksheet firstSheet = workbook.Worksheets[0];

            // Demonstrate that the worksheet is accessible (write a value to cell A1)
            firstSheet.Cells["A1"].PutValue("Accessed by index");

            // Save the workbook to verify the changes
            workbook.Save("WorksheetByIndex.xlsx");
        }
    }
}
