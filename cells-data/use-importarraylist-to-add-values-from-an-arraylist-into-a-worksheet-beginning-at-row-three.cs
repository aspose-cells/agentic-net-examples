using System;
using System.Collections;
using Aspose.Cells;

class ImportArrayListExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Prepare an ArrayList with sample data
        ArrayList data = new ArrayList();
        data.Add("Alice");
        data.Add(28);
        data.Add("Engineer");

        // Import the ArrayList into the worksheet
        // Row index 2 corresponds to the third row (row 3 in Excel)
        // Column index 0 corresponds to column A
        // false => import horizontally (across columns)
        cells.ImportArrayList(data, 2, 0, false);

        // Save the workbook to a file
        workbook.Save("ImportArrayListRow3.xlsx");
    }
}