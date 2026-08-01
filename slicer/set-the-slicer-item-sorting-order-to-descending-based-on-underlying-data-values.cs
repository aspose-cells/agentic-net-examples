// Title: C# – Sort Excel Slicer Items Descending with Aspose.Cells
// Description: Shows how to open an .xlsx workbook, locate the first slicer on a worksheet, apply SortOrder.Descending to its items, and write the changes back to a new file.
// Keywords: Aspose.Cells | C# slicer sorting | Excel slicer descending | SortOrder.Descending | programmatic slicer order | dashboard automation .NET | Aspose.Cells Slicer API
// Common Searches: Aspose.Cells set slicer sort order descending | C# change Excel slicer item order | How to sort slicer values descending using Aspose | Programmatically reorder slicer items .NET | Excel slicer descending order code example
// Developer Intent: Configure a slicer so its entries are displayed from the highest to the lowest value derived from the source data.
// Use Cases: Generate a sales dashboard where the region slicer lists top‑performing areas first. | Standardize slicer ordering across multiple reports created in an automated pipeline. | Prepare a workbook for distribution, ensuring users see the most significant items at the top of each slicer.
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, finds a specific slicer, and sets its SortOrder to descending while safely handling missing slicers. | Create a snippet that toggles a slicer's SortOrder between ascending and descending based on a runtime flag in a .NET application. | Provide error‑handling patterns for applying SortOrder.Descending when the target worksheet contains no slicers or the slicer name is unknown.

using Aspose.Cells;
using Aspose.Cells.Slicers;

// Shows how to open an .xlsx workbook, locate the first slicer on a worksheet, apply SortOrder.Descending to its items, and write the changes back to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook that contains a slicer
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Verify that the worksheet has at least one slicer
        if (sheet.Slicers.Count > 0)
        {
            // Retrieve the first slicer
            Slicer slicer = sheet.Slicers[0];

            // Set the slicer items to be sorted in descending order
            slicer.SortOrderType = SortOrder.Descending;
        }

        // Save the workbook with the updated slicer sorting
        workbook.Save("output.xlsx");
    }
}
