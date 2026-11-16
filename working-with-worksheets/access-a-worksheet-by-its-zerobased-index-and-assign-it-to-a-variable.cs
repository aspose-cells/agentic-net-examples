// Title: Access a worksheet by its zero‑based index and assign it to a variable with Aspose.Cells for C#
// Common Searches: Aspose.Cells C# get worksheet by index from workbook | retrieve worksheet using zero based index Aspose.Cells C# example | how to assign a worksheet object to a variable in Aspose.Cells | C# Aspose.Cells access first sheet by position | using Worksheets collection index to reference sheet in Aspose.Cells
// Tags: workbook Worksheets collection access Aspose.Cells C# | assign worksheet variable Aspose.Cells C# | modify worksheet name property Aspose.Cells | save workbook after sheet changes Aspose.Cells | zero‑based sheet indexing Aspose.Cells C#

using System;
using Aspose.Cells;

// The example creates (or loads) a Workbook, accesses the first worksheet via its zero‑based index using the Worksheets collection, assigns the Worksheet to a variable, optionally changes its Name, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates an empty workbook

        // Access the first worksheet by its zero‑based index (0)
        int sheetIndex = 0;
        Worksheet worksheet = workbook.Worksheets[sheetIndex];

        // Example: set the name of the accessed worksheet
        worksheet.Name = "FirstSheet";

        // Save the workbook to verify changes (optional)
        workbook.Save("Result.xlsx");
    }
}
