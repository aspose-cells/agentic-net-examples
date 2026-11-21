// Title: How to set a custom height for the header row and auto‑fit subsequent rows using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set specific row height for header and auto‑fit other rows | Aspose.Cells .NET auto‑fit rows starting from second row | How to change header row height in an Excel file with Aspose.Cells C# | Aspose.Cells set row height in points and auto‑size remaining rows | Programmatically adjust Excel row heights with Aspose.Cells in C#
// Tags: custom row height Aspose.Cells C# | exclude first row from AutoFitRows Aspose.Cells | header row formatting Excel Aspose.Cells | auto‑size data rows Aspose.Cells .NET | adjust Excel row heights programmatically C#

using Aspose.Cells;

// Sets the first worksheet row height to 30 points, adds sample data, auto‑fits rows 2 through the last populated row, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Set header row (row index 0) height to 30 points for visual emphasis
        sheet.Cells.SetRowHeight(0, 30);

        // Example data to demonstrate auto‑fit (can be replaced with actual data)
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue("Item 1");
        sheet.Cells["B2"].PutValue(123);
        sheet.Cells["A3"].PutValue("Item 2");
        sheet.Cells["B3"].PutValue(456);

        // Determine the last row with data
        int lastRow = sheet.Cells.MaxDataRow;

        // Auto‑fit all rows except the header row
        if (lastRow > 0)
        {
            sheet.AutoFitRows(1, lastRow);
        }

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
