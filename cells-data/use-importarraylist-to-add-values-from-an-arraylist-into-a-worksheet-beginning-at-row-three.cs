using System;
using System.Collections;
using Aspose.Cells;

class ImportArrayListExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the Cells collection of the first worksheet
        Cells cells = workbook.Worksheets[0].Cells;

        // Prepare an ArrayList with sample data
        ArrayList data = new ArrayList();
        data.Add("Alice");
        data.Add(28);
        data.Add(DateTime.Now);

        // Import the ArrayList into the worksheet.
        // Start at row 3 (zero‑based index 2), column 1 (index 0), horizontally (isVertical = false)
        cells.ImportArrayList(data, 2, 0, false);

        // Save the workbook to a file
        workbook.Save("ImportArrayListRow3.xlsx");
    }
}