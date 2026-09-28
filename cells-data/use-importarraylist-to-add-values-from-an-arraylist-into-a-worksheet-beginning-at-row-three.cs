// Title: Import a .NET ArrayList into an Aspose.Cells worksheet beginning at row 3 (C#)
// AI Prompts: Write C# code that calls Cells.ImportArrayList to insert an ArrayList horizontally starting at row index 2, column 0 in an Aspose.Cells workbook. | Show how to save the workbook to "ImportArrayListRow3.xlsx" after importing the ArrayList. | Explain how to switch the ImportArrayList orientation to vertical if the data should fill down rows instead of across columns.
// Common Searches: Aspose.Cells C# import ArrayList starting at third row example | How to write a .NET ArrayList to a specific row in Excel using Aspose.Cells | Cells.ImportArrayList horizontal import row index 2 C# tutorial | Saving workbook after ImportArrayList with Aspose.Cells | Zero‑based row index usage in Aspose.Cells ImportArrayList
// Tags: ImportArrayList horizontal import Aspose.Cells C# | populate worksheet from ArrayList Aspose.Cells | write data to third row Aspose.Cells | Aspose.Cells save workbook after import | zero‑based row index ImportArrayList

using System;
using System.Collections;
using Aspose.Cells;

// The example creates a new Workbook, builds an ArrayList with sample values, and uses Cells.ImportArrayList to write the list horizontally beginning at row index 2 (the third row) and column A. After the import, the workbook is saved as ImportArrayListRow3.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the cells collection of the first worksheet
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Prepare an ArrayList with sample values
        ArrayList data = new ArrayList();
        data.Add("Alice");
        data.Add(28);
        data.Add("Engineer");

        // Import the ArrayList starting at row 3 (zero‑based index 2), column A (index 0)
        // false = horizontal import (values go across columns)
        cells.ImportArrayList(data, 2, 0, false);

        // Save the workbook to a file
        workbook.Save("ImportArrayListRow3.xlsx");
    }
}
