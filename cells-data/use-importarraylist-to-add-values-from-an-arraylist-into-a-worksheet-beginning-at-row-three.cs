// Title: How to import a System.Collections.ArrayList into an Aspose.Cells worksheet starting at row 3 in C#
// AI Prompts: Write C# code that uses Worksheet.Cells.ImportArrayList to place each element of an ArrayList into successive rows beginning at the third worksheet row. | Show an example of offsetting the import location to row index 2 and column index 0 when importing a list with Aspose.Cells. | Demonstrate saving the workbook after importing data from an ArrayList with a specific start row using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# import ArrayList starting at row 3 example | Worksheet.Cells.ImportArrayList offset row and column parameters | How to write a vertical list from ArrayList into Excel with Aspose.Cells | C# Aspose.Cells import collection into specific worksheet cell range
// Tags: ImportArrayList row offset Aspose.Cells | vertical list import Aspose.Cells C# | Worksheet.Cells.ImportArrayList example | populate worksheet from ArrayList Aspose.Cells | start import at specific row Aspose.Cells

using Aspose.Cells;
using System.Collections;

// Creates a new workbook, builds an ArrayList with a header and names, and uses Worksheet.Cells.ImportArrayList to insert the items vertically starting at row 3 (zero‑based index 2) in the first column, then saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Prepare the ArrayList with data to import
        ArrayList data = new ArrayList();
        data.Add("Header");
        data.Add("Alice");
        data.Add("Bob");
        data.Add("Charlie");

        // Import the ArrayList starting at row 3 (zero‑based index 2), column 1 (index 0)
        // Set isVertical = true to place each item in a new row
        sheet.Cells.ImportArrayList(data, 2, 0, true);

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
