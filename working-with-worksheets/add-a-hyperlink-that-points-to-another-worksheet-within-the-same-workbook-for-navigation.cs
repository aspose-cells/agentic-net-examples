// Title: Add an internal worksheet hyperlink that navigates to another sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells to create a workbook, add two worksheets, and insert a hyperlink in Sheet1!A1 that opens Sheet2!A1. | Write a C# example that uses Aspose.Cells to add an internal hyperlink between worksheets, using the "SheetName!CellAddress" address format.
// Common Searches: asp.net cells c# add hyperlink to another worksheet in same workbook | how to create internal sheet link using Aspose.Cells .NET | Aspose.Cells hyperlink address format for same workbook | C# example linking Sheet1 cell to Sheet2 cell with Aspose.Cells | navigate between worksheets programmatically Aspose.Cells
// Tags: Aspose.Cells add internal worksheet hyperlink | C# Aspose.Cells hyperlink address SheetName!CellAddress | Aspose.Cells create multi-sheet workbook | Aspose.Cells worksheet navigation link | Aspose.Cells save workbook as XLSX

using Aspose.Cells;

// The example creates a new workbook, adds two worksheets named Sheet1 and Sheet2, writes "Go to Sheet2" in Sheet1!A1, adds an internal hyperlink in that cell that points to Sheet2!A1 using the "SheetName!CellAddress" format, and saves the file as HyperlinkNavigation.xlsx.
class HyperlinkExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (default)
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";

        // Add a second worksheet
        int sheet2Index = workbook.Worksheets.Add();
        Worksheet sheet2 = workbook.Worksheets[sheet2Index];
        sheet2.Name = "Sheet2";

        // Set display text in Sheet1!A1
        Cell cell = sheet1.Cells["A1"];
        cell.PutValue("Go to Sheet2");

        // Add a hyperlink in Sheet1!A1 that points to Sheet2!A1
        // The hyperlink address for an internal sheet reference uses the format "SheetName!CellAddress"
        sheet1.Hyperlinks.Add(0, 0, 1, 1, "Sheet2!A1");

        // Save the workbook to a file
        workbook.Save("HyperlinkNavigation.xlsx");
    }
}
