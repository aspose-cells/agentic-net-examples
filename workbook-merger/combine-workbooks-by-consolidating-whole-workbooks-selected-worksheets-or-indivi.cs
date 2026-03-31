using System;
using Aspose.Cells;

class CombineWorkbooksDemo
{
    static void Main()
    {
        // Load the first source workbook from disk
        Workbook sourceWorkbook1 = new Workbook("Source1.xlsx");

        // Load the second source workbook from disk
        Workbook sourceWorkbook2 = new Workbook("Source2.xlsx");

        // Create an empty destination workbook that will hold the combined result
        Workbook destinationWorkbook = new Workbook();

        // Combine the entire first workbook into the destination workbook
        destinationWorkbook.Combine(sourceWorkbook1);

        // Combine the entire second workbook into the destination workbook
        destinationWorkbook.Combine(sourceWorkbook2);

        // OPTIONAL: Copy a specific worksheet from one of the source workbooks
        // Here we copy the first worksheet of sourceWorkbook1 as a separate sheet
        int copiedSheetIndex = destinationWorkbook.Worksheets.AddCopy(sourceWorkbook1.Worksheets[0].Name);
        // Rename the copied sheet for clarity
        destinationWorkbook.Worksheets[copiedSheetIndex].Name = "Copied_From_Source1";

        // Save the combined workbook to a new file
        destinationWorkbook.Save("CombinedResult.xlsx", SaveFormat.Xlsx);
    }
}