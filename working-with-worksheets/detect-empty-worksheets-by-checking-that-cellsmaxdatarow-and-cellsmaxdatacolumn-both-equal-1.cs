// Title: Detect Empty Worksheets in Aspose.Cells for .NET (C#) Using MaxDataRow & MaxDataColumn
// Description: Load an Excel workbook with Aspose.Cells, iterate its worksheets, and determine emptiness by confirming both Cells.MaxDataRow and Cells.MaxDataColumn return -1. The sample prints the result and saves the file unchanged.
// Keywords: Aspose.Cells | C# | .NET | empty worksheet detection | MaxDataRow | MaxDataColumn | -1 empty sheet | Excel automation | worksheet data check | detect empty sheet
// Common Searches: Aspose.Cells check if worksheet is empty | MaxDataRow -1 meaning in Aspose.Cells | C# detect empty Excel sheet with Aspose | list empty worksheets in a workbook using Aspose.Cells | identify blank worksheets programmatically .NET
// Developer Intent: Find worksheets that contain no data by evaluating MaxDataRow and MaxDataColumn values.
// Use Cases: Log names of empty sheets before data extraction or transformation. | Skip conversion or export of blank worksheets when generating PDFs, images, or other formats. | Validate required worksheets are populated and flag any that are completely empty.
// AI Prompts: Write a C# method that returns a list of worksheet names that are empty using Aspose.Cells MaxDataRow and MaxDataColumn checks. | Provide code to remove all empty worksheets from a workbook based on MaxDataRow and MaxDataColumn values. | Explain how MaxDataRow and MaxDataColumn behave when a worksheet has only formatting but no cell values.

using System;
using Aspose.Cells;

// Load an Excel workbook with Aspose.Cells, iterate its worksheets, and determine emptiness by confirming both Cells.MaxDataRow and Cells.MaxDataColumn return -1. The sample prints the result and saves the file unchanged.
class DetectEmptyWorksheets
{
    static void Main()
    {
        // Load an existing workbook (provide the correct path to your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // A worksheet is considered empty when both MaxDataRow and MaxDataColumn are -1
            if (sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" is empty.");
            }
            else
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains data.");
            }
        }

        // Save the workbook (no modifications made, just demonstrating save usage)
        workbook.Save("output.xlsx");
    }
}
