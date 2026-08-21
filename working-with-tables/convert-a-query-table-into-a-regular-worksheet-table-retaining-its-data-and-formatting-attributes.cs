// Title: Convert a QueryTable‑linked ListObject to a regular range with Aspose.Cells for .NET
// Description: Loads an Excel workbook, finds the first ListObject that has an associated QueryTable, calls ListObject.ConvertToRange() to remove the table structure while preserving all cell values and formatting, and saves the result. Demonstrates the Aspose.Cells .NET API for flattening query tables.
// Keywords: Aspose.Cells ConvertToRange | C# convert query table to range | Aspose.Cells ListObject QueryTable | remove table structure keep formatting | Aspose.Cells .NET example | Excel query table to range | ListObject ConvertToRange method
// Common Searches: convert query table to range Aspose.Cells | Aspose.Cells remove ListObject but keep data | C# find ListObject with QueryTable | Aspose.Cells ConvertToRange usage | flatten Excel query table with Aspose
// Developer Intent: Turn a query‑linked Excel table into a plain worksheet range without losing data or formatting.
// Use Cases: Flatten imported query data for downstream calculations. | Make workbook compatible with tools that do not support Excel tables. | Simplify cell references after data import.
// AI Prompts: Write C# code using Aspose.Cells to locate a ListObject that has a QueryTable and convert it to a regular range. | Explain what happens to the QueryTable object after calling ListObject.ConvertToRange in Aspose.Cells. | Provide a step‑by‑step guide to load a workbook, convert the first query‑linked table to a range, and save the file using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Tables;

// Loads an Excel workbook, finds the first ListObject that has an associated QueryTable, calls ListObject.ConvertToRange() to remove the table structure while preserving all cell values and formatting, and saves the result. Demonstrates the Aspose.Cells .NET API for flattening query tables.
class ConvertQueryTableToRange
{
    static void Main()
    {
        // Load the workbook that contains the query table
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Locate the ListObject (table) that is linked to a QueryTable
        ListObject listObjectWithQuery = null;
        foreach (ListObject lo in sheet.ListObjects)
        {
            if (lo.QueryTable != null)
            {
                listObjectWithQuery = lo;
                break;
            }
        }

        // If such a ListObject is found, convert it to a normal range
        if (listObjectWithQuery != null)
        {
            // This removes the table structure but keeps the data and formatting
            listObjectWithQuery.ConvertToRange();
        }

        // Save the workbook with the query table now converted to a regular range
        workbook.Save("output.xlsx");
    }
}
