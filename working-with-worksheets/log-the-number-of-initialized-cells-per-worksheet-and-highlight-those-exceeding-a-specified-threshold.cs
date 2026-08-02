// Title: Log worksheet cell count and highlight tabs exceeding a threshold with Aspose.Cells for .NET
// Description: Creates a workbook, fills two worksheets, uses Cells.CountLarge to obtain the number of instantiated cells per sheet, writes the counts to the console, and changes the worksheet tab color to LightCoral when the count is above a defined limit before saving the file.
// Keywords: Aspose.Cells | C# | Cells.CountLarge | worksheet tab color | threshold cell count | log cell count | initialize cells | Excel automation
// Common Searches: Aspose.Cells get number of initialized cells | change worksheet tab color based on cell count Aspose | Cells.CountLarge example C# | highlight large worksheets Aspose.Cells | log worksheet cell count .NET
// Developer Intent: Identify how many cells are instantiated in each worksheet and automatically flag sheets that exceed a specified count.
// Use Cases: Produce a console report of instantiated cells per worksheet for workbook size monitoring. | Visually flag worksheets with excessive data by setting their tab color. | Save the workbook after applying visual cues to help end‑users locate large sheets.
// AI Prompts: Generate C# code that iterates all worksheets, logs Cells.CountLarge, and sets TabColor when the count surpasses a given threshold using Aspose.Cells. | Explain the difference between Cells.CountLarge and Cells.MaxDataColumn and when to prefer CountLarge for memory‑usage checks. | Show how to export the logged worksheet cell counts to a CSV file instead of console output.

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a workbook, fills two worksheets, uses Cells.CountLarge to obtain the number of instantiated cells per sheet, writes the counts to the console, and changes the worksheet tab color to LightCoral when the count is above a defined limit before saving the file.
class LogAndHighlightCells
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Populate first worksheet with a few cells
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";
        ws1.Cells["A1"].PutValue("Hello");
        ws1.Cells["B2"].PutValue(123);
        ws1.Cells["C3"].PutValue(DateTime.Now);

        // Add a second worksheet and fill many cells to exceed the threshold
        int secondIndex = workbook.Worksheets.Add();
        Worksheet ws2 = workbook.Worksheets[secondIndex];
        ws2.Name = "Sheet2";
        for (int i = 0; i < 500; i++)
        {
            ws2.Cells[i, 0].PutValue(i);
        }

        // Threshold for highlighting
        long threshold = 200;

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Use Cells.CountLarge to get the number of instantiated cells (rule)
            long cellCount = sheet.Cells.CountLarge;

            // Log the count
            Console.WriteLine($"Worksheet \"{sheet.Name}\" has {cellCount} initialized cells.");

            // Highlight worksheet tab if count exceeds threshold
            if (cellCount > threshold)
            {
                sheet.TabColor = Color.LightCoral; // Highlight with a noticeable color
                Console.WriteLine($"  -> Exceeds threshold ({threshold}); tab color set to LightCoral.");
            }
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("LogAndHighlight.xlsx");
    }
}
