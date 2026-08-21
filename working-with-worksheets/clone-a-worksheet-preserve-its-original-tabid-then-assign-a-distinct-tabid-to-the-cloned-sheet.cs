// Title: Clone a worksheet in Aspose.Cells (C#) while preserving the original TabId and assigning a new TabId
// Description: Demonstrates how to create a workbook, set a custom TabId on the source sheet, duplicate the sheet with Workbook.Worksheets.AddCopy, rename the clone, assign a different TabId, and save the result as an Excel file.
// Keywords: Aspose.Cells | C# | clone worksheet | TabId property | AddCopy | preserve TabId | set TabId on copy | duplicate sheet | .NET Excel API | worksheet copy example
// Common Searches: How to clone a worksheet in Aspose.Cells and keep its TabId | Assign a new TabId to a copied sheet using Aspose.Cells for .NET | Aspose.Cells AddCopy preserve original TabId | Set TabId for cloned worksheet in C# | Copy worksheet with custom TabId Aspose.Cells
// Developer Intent: The developer needs to duplicate an existing worksheet, retain the original sheet's TabId, and give the cloned sheet a separate TabId for later identification.
// Use Cases: Create a master template sheet, clone it for each department, and give each clone a unique TabId to simplify navigation in large workbooks. | Generate periodic report sheets from a base layout while keeping the source TabId unchanged for audit tracking. | Automate the production of multiple worksheets from a single design, assigning sequential TabIds to support API‑driven sheet selection.
// AI Prompts: Provide C# code that clones a worksheet in Aspose.Cells, preserves the original TabId, and sets a different TabId on the clone. | Explain the behavior of the TabId property when using Workbook.Worksheets.AddCopy to duplicate a sheet. | Show an example that copies several worksheets from a template and assigns distinct TabIds to each new sheet.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, set a custom TabId on the source sheet, duplicate the sheet with Workbook.Worksheets.AddCopy, rename the clone, assign a different TabId, and save the result as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet, give it a name and set its TabId
        Worksheet original = workbook.Worksheets[0];
        original.Name = "Original";
        original.TabId = 100; // preserve original TabId

        // Add some sample data
        original.Cells["A1"].PutValue("Original sheet");

        // Clone the worksheet using AddCopy (by name)
        int clonedIndex = workbook.Worksheets.AddCopy("Original");
        Worksheet cloned = workbook.Worksheets[clonedIndex];
        cloned.Name = "Cloned";

        // Assign a distinct TabId to the cloned sheet
        cloned.TabId = 200;

        // Save the workbook
        workbook.Save("ClonedWorksheetTabIdDemo.xlsx");
    }
}
