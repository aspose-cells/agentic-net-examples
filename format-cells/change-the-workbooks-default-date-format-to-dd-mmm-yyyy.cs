// Title: Set default workbook date format to dd‑MMM‑yyyy using Aspose.Cells for .NET
// Description: Creates a new Workbook, modifies its DefaultStyle.Custom property to "dd-MMM-yyyy", optionally applies the style to a sample cell, and saves the file. This changes the global date display for all new cells in the workbook.
// Keywords: Aspose.Cells default date format | C# set workbook date style | custom date format dd-MMM-yyyy | global date format Aspose.Cells | .NET workbook default style | change default date pattern
// Common Searches: Aspose.Cells change default date format | set workbook default style to dd-MMM-yyyy .NET | global date format for all cells Aspose.Cells | how to apply custom date format to new workbook | default date pattern Aspose.Cells C#
// Developer Intent: Change the workbook’s default date format to dd‑MMM‑yyyy so every new date cell uses this pattern automatically.
// Use Cases: Generate a template where all dates appear as dd‑MMM‑yyyy without per‑cell styling. | Create financial reports that enforce a consistent date representation across sheets. | Distribute a shared workbook that guarantees downstream users see dates in the required format.
// AI Prompts: Provide C# code that sets Aspose.Cells Workbook.DefaultStyle.Custom to "dd-MMM-yyyy" and demonstrates the result in a cell. | Explain step‑by‑step how to change the global date format of a workbook using Aspose.Cells for .NET. | Show how to update existing worksheets after modifying the default date style in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsDateFormatDemo
{
    // Creates a new Workbook, modifies its DefaultStyle.Custom property to "dd-MMM-yyyy", optionally applies the style to a sample cell, and saves the file. This changes the global date display for all new cells in the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the workbook's default style and set a custom date format
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Custom = "dd-mmm-yyyy";
            workbook.DefaultStyle = defaultStyle; // apply the modified default style

            // Demonstrate the default format by inserting a date into a cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(DateTime.Now);
            cell.SetStyle(defaultStyle); // optional: apply explicitly to the cell

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DefaultDateFormat.xlsx");
        }
    }
}
