// Title: C# – Import an ArrayList into an Aspose.Cells worksheet starting at row 3 (vertical orientation)
// Description: Creates a new Workbook, builds an ArrayList with a string, an integer, another string and a DateTime, then uses Cells.ImportArrayList to write each item vertically beginning at the third worksheet row (zero‑based index 2) in column A and saves the file as ImportArrayListDemo.xlsx.
// Keywords: Aspose.Cells | ImportArrayList | C# | Excel export | ArrayList to worksheet | vertical import | start row 3 | mixed data types | save workbook | cells.ImportArrayList
// Common Searches: Aspose.Cells ImportArrayList C# example | How to import an ArrayList into Excel at a specific row | Import vertical list into worksheet using Aspose.Cells | Start importing data at row 3 Aspose.Cells | Import mixed type collection to Excel C#
// Developer Intent: Insert each element of an ArrayList into consecutive rows of a worksheet, beginning with row 3, column A, using vertical orientation.
// Use Cases: Fill a pre‑designed report template with employee records starting at row 3 without overwriting headers. | Append runtime‑generated data to a designated section of an existing Excel sheet. | Export a mixed‑type collection (text, numbers, dates) where each value occupies its own row for downstream processing.
// AI Prompts: Show how to import a List<T> into an Aspose.Cells worksheet starting at row 5 horizontally. | Explain how to import a DataTable into a worksheet with a header row using ImportDataTable. | Provide code to import a two‑dimensional array into a worksheet at a custom start cell.

using System;
using System.Collections;
using Aspose.Cells;

// Creates a new Workbook, builds an ArrayList with a string, an integer, another string and a DateTime, then uses Cells.ImportArrayList to write each item vertically beginning at the third worksheet row (zero‑based index 2) in column A and saves the file as ImportArrayListDemo.xlsx.
class ImportArrayListExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet's cells collection
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Prepare the data to be imported
        ArrayList data = new ArrayList();
        data.Add("Alice");
        data.Add(28);
        data.Add("Engineer");
        data.Add(DateTime.Now);

        // Import the ArrayList starting at row 3 (zero‑based index 2), column A (index 0)
        // Set isVertical to true to place each item in a new row vertically
        cells.ImportArrayList(data, firstRow: 2, firstColumn: 0, isVertical: true);

        // Save the workbook
        workbook.Save("ImportArrayListDemo.xlsx");
    }
}
