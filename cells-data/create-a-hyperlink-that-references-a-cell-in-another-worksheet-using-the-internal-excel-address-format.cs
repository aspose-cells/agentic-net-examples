using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the default first worksheet and rename it
        Worksheet mainSheet = workbook.Worksheets[0];
        mainSheet.Name = "Main";

        // Add a second worksheet and rename it
        int targetSheetIndex = workbook.Worksheets.Add();
        Worksheet targetSheet = workbook.Worksheets[targetSheetIndex];
        targetSheet.Name = "Target";

        // Put a sample value in the target cell (A1 of the second sheet)
        targetSheet.Cells["A1"].PutValue("Hello from Target!");

        // Add a hyperlink in Main!B2 that points to Target!A1 using internal Excel address format
        // The Add method parameters: start cell, rows, columns, address
        mainSheet.Hyperlinks.Add("B2", 1, 1, "Target!A1");

        // Optionally set the display text of the hyperlink
        Hyperlink hyperlink = mainSheet.Hyperlinks[0];
        hyperlink.TextToDisplay = "Go to Target A1";

        // Save the workbook
        workbook.Save("HyperlinkToOtherSheet.xlsx");
    }
}