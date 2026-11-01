// Title: Add an internal hyperlink from Sheet1!B2 to Sheet2!A1 using Aspose.Cells for .NET
// AI Prompts: Insert a hyperlink in cell B2 of Sheet1 that points to cell A1 on Sheet2 using the Aspose.Cells Hyperlinks.Add method with an internal Excel address. | Create the internal address string for a worksheet cell (e.g., 'Sheet2'!A1) and apply it to a hyperlink in a .NET workbook. | Assign custom display text to a cell containing an internal worksheet hyperlink and save the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# create hyperlink to another worksheet cell using internal address | how to use Hyperlinks.Add for internal sheet reference in Aspose.Cells | C# example linking Sheet1 cell to Sheet2 A1 with Aspose.Cells | internal Excel address format for Aspose.Cells hyperlink | set display text for internal worksheet hyperlink Aspose.Cells .NET
// Tags: Aspose.Cells hyperlink API | C# internal worksheet hyperlink Aspose.Cells | Excel hyperlink address syntax Aspose.Cells | set hyperlink caption Aspose.Cells | save workbook as .xlsx Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook with two worksheets, adds an internal hyperlink in Sheet1!B2 that points to Sheet2!A1 using Aspose.Cells, sets custom display text, and saves the file as HyperlinkInternal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and rename it
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet and rename it
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Sheet2";

        // Put some data in the target cell on Sheet2
        sheet2.Cells["A1"].PutValue("Target Cell");

        // Define the internal Excel address that points to Sheet2!A1
        // Use single quotes around the sheet name in case it contains spaces or special characters
        string internalAddress = "'Sheet2'!A1";

        // Add a hyperlink in Sheet1!B2 that points to the internal address
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hyperlink address
        sheet1.Hyperlinks.Add(1, 1, 1, 1, internalAddress);

        // Set the display text for the hyperlink cell
        sheet1.Cells["B2"].PutValue("Go to Sheet2!A1");

        // Save the workbook to a file
        workbook.Save("HyperlinkInternal.xlsx");
    }
}
