// Title: Enumerate all worksheet names in an Excel workbook using Aspose.Cells for .NET (C#)
// Description: This example shows how to open an existing Excel file with Aspose.Cells, loop through the Workbook.Worksheets collection, and print each Worksheet.Name to the console. No changes are saved, making it ideal for read‑only scenarios.
// Keywords: Aspose.Cells list worksheets | C# read Excel sheet names | enumerate workbook worksheets .NET | display worksheet names console | Aspose.Cells get sheet titles
// Common Searches: how to get all sheet names from an Excel file using Aspose.Cells | C# loop through worksheets in a workbook | Aspose.Cells print worksheet names | read Excel worksheet titles without saving
// Developer Intent: Obtain and show the names of every worksheet contained in a loaded workbook.
// Use Cases: Verify required worksheets are present before data processing | Create an audit log of all sheet names in a workbook | Dynamically select a worksheet by name after listing available sheets
// AI Prompts: Generate C# code with Aspose.Cells that lists all worksheet names and includes try‑catch handling for missing or corrupted files. | Provide an example that filters worksheets whose names begin with a given prefix while enumerating the workbook. | Show how to write the list of worksheet names to a text file instead of the console using Aspose.Cells.

using System;
using Aspose.Cells;

// This example shows how to open an existing Excel file with Aspose.Cells, loop through the Workbook.Worksheets collection, and print each Worksheet.Name to the console. No changes are saved, making it ideal for read‑only scenarios.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enumerate all worksheets and output their names
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine(sheet.Name);
        }

        // No need to save if only reading; uncomment if you modify and want to save
        // workbook.Save("output.xlsx");
    }
}
