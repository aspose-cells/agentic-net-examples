// Title: Set Aspose.Cells slicer to non‑printable in C# (.NET)
// Description: Creates a workbook, adds a table, inserts a slicer linked to the first column, disables its printing by setting Shape.IsPrintable = false, and saves the file so the slicer is hidden in printed or PDF output.
// Keywords: Aspose.Cells | C# slicer | IsPrintable | non printable slicer | hide slicer from print | prevent slicer printing | Excel slicer Aspose | Aspose.Cells .NET example
// Common Searches: Aspose.Cells make slicer non printable | C# hide slicer when printing Excel | Set slicer Shape.IsPrintable false | Exclude slicer from PDF export Aspose | Aspose.Cells slicer print settings
// Developer Intent: Disable printing of a slicer object in an Excel workbook.
// Use Cases: Design interactive dashboards where slicers are visible on screen but omitted from hard‑copy reports. | Automate generation of printable Excel or PDF files that contain slicers for filtering but should not appear on the final page. | Create templates that add slicers for user interaction and then suppress their print output before distribution.
// AI Prompts: Provide C# code using Aspose.Cells to add a slicer to a table and set its Shape.IsPrintable property to false. | Explain how the IsPrintable flag affects Excel, PDF, and printer output for slicers in Aspose.Cells. | Show how to test that a slicer does not appear in the printed version of a workbook created with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

// Creates a workbook, adds a table, inserts a slicer linked to the first column, disables its printing by setting Shape.IsPrintable = false, and saves the file so the slicer is hidden in printed or PDF output.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["A2"].Value = "A";
        worksheet.Cells["A3"].Value = "B";
        worksheet.Cells["B1"].Value = "Amount";
        worksheet.Cells["B2"].Value = 100;
        worksheet.Cells["B3"].Value = 200;

        // Convert the range into a table (ListObject)
        ListObject table = worksheet.ListObjects[worksheet.ListObjects.Add("A1", "B3", true)];

        // Add a slicer linked to the first column of the table
        int slicerIndex = worksheet.Slicers.Add(table, 0, "D1");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Mark the slicer as non‑printable so it will not appear in printed output
        slicer.Shape.IsPrintable = false;

        // Save the workbook
        workbook.Save("SlicerNonPrintable.xlsx");
    }
}
