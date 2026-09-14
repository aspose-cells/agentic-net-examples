// Title: Create an internal worksheet hyperlink that points to a cell on another sheet using Aspose.Cells for .NET
// AI Prompts: Insert a hyperlink in Sheet1!A1 that navigates to Sheet2!A1 using Aspose.Cells' internal address format. | Set the TextToDisplay property of the worksheet hyperlink and save the workbook as a .xlsx file. | Add a second worksheet named "Sheet2", write a value to A1, and link to it from the first sheet.
// Common Searches: Aspose.Cells how to add a hyperlink that points to a cell on another worksheet | C# create internal Excel address format hyperlink with Aspose.Cells | Set display text for a worksheet hyperlink using Aspose.Cells .NET | Programmatically link Sheet1 to Sheet2 cell A1 in Aspose.Cells | Add hyperlink between worksheets in a workbook using Aspose.Cells for .NET
// Tags: Aspose.Cells internal worksheet hyperlink | Aspose.Cells hyperlink to another sheet | Aspose.Cells set hyperlink display text | Aspose.Cells add second worksheet programmatically | Aspose.Cells C# hyperlink address format

using System;
using Aspose.Cells;

// The example creates a workbook, adds a second worksheet named "Sheet2", writes "Target Cell" to Sheet2!A1, inserts a hyperlink in Sheet1!A1 that points to Sheet2!A1 using the internal Excel address format, customizes the hyperlink's display text, and saves the file as HyperlinkToAnotherSheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a second worksheet and give it a name
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Sheet2";

        // Put a value in the target cell of the second worksheet
        sheet2.Cells["A1"].PutValue("Target Cell");

        // Get reference to the first worksheet (default sheet)
        Worksheet sheet1 = workbook.Worksheets[0];

        // Add a hyperlink in cell A1 of the first worksheet that points to Sheet2!A1
        // Using the internal Excel address format "Sheet2!A1"
        int hyperlinkIndex = sheet1.Hyperlinks.Add("A1", 1, 1, "Sheet2!A1");

        // Optionally set the display text for the hyperlink
        Hyperlink hyperlink = sheet1.Hyperlinks[hyperlinkIndex];
        hyperlink.TextToDisplay = "Go to Sheet2 A1";

        // Save the workbook
        workbook.Save("HyperlinkToAnotherSheet.xlsx");
    }
}
