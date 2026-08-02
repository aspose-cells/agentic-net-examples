// Title: Enumerate all worksheet names in an Excel file using Aspose.Cells for .NET (C#)
// Description: This C# example shows how to open an existing .xlsx workbook with Aspose.Cells, access the Workbook.Worksheets collection, iterate through each Worksheet object, and output its Name property to the console.
// Keywords: Aspose.Cells list worksheets | C# read Excel sheet names | Aspose.Cells get worksheet collection | print worksheet names .NET | enumerate Excel sheets C#
// Common Searches: aspocells get worksheet names c# | how to list all sheets in an Excel workbook using Aspose.Cells | C# loop through worksheets Aspose.Cells example | retrieve sheet names from .xlsx with Aspose.Cells | display Excel worksheet names console Aspose
// Developer Intent: Show every sheet title contained in a loaded workbook.
// Use Cases: Confirm required tabs exist before extracting data | Log workbook structure for debugging or audit trails | Select a sheet dynamically based on its retrieved name | Generate documentation of a workbook’s layout
// AI Prompts: Write a C# function that returns a List<string> of worksheet names from a given Excel file using Aspose.Cells. | Adapt the example to write sheet names to a text file instead of the console. | Provide code that filters worksheet names starting with a specific prefix (e.g., "Report_") using Aspose.Cells. | Explain how to include hidden worksheets when enumerating sheet names.

using System;
using Aspose.Cells;

// This C# example shows how to open an existing .xlsx workbook with Aspose.Cells, access the Workbook.Worksheets collection, iterate through each Worksheet object, and output its Name property to the console.
class Program
{
    static void Main(string[] args)
    {
        // Path to the existing workbook file
        string workbookPath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(workbookPath);

        // Get the collection of worksheets
        WorksheetCollection worksheets = workbook.Worksheets;

        // Enumerate each worksheet and output its name
        for (int i = 0; i < worksheets.Count; i++)
        {
            Worksheet sheet = worksheets[i];
            Console.WriteLine(sheet.Name);
        }
    }
}
