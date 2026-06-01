using System;
using Aspose.Cells;

class HyperlinkToAnotherWorksheet
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet and name it "Sheet2"
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Sheet2";

        // Put a value in the target cell on Sheet2
        sheet2.Cells["A1"].PutValue("Target Cell");

        // Get the first worksheet (default name "Sheet1")
        Worksheet sheet1 = workbook.Worksheets[0];

        // Add a hyperlink in cell A1 of Sheet1 that points to Sheet2!A1
        // Using the internal Excel address format "Sheet2!A1"
        sheet1.Hyperlinks.Add("A1", 1, 1, "Sheet2!A1");

        // Optionally set the display text for the hyperlink
        sheet1.Hyperlinks[0].TextToDisplay = "Go to Sheet2 A1";

        // Save the workbook
        workbook.Save("HyperlinkToAnotherWorksheet.xlsx");
    }
}