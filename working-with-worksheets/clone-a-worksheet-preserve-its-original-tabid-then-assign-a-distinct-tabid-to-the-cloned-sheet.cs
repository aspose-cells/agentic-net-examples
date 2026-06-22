using System;
using Aspose.Cells;

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

        // Add sample data to the original sheet
        original.Cells["A1"].PutValue("Data in original sheet");

        // Clone the worksheet using AddCopy (by index)
        int clonedIndex = workbook.Worksheets.AddCopy(original.Index);
        Worksheet cloned = workbook.Worksheets[clonedIndex];
        cloned.Name = "Cloned";

        // Assign a distinct TabId to the cloned sheet
        cloned.TabId = original.TabId + 1; // ensure different TabId

        // Save the workbook
        workbook.Save("ClonedWorksheet.xlsx");
    }
}