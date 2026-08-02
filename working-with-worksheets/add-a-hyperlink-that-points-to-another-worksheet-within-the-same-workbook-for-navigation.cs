// Title: Add internal worksheet navigation hyperlinks with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds two sheets (Sheet1 and Sheet2), writes link text in cell A1 of each sheet, and uses the Hyperlinks.Add method to insert clickable links that jump to the opposite sheet's A1 cell. Custom display text is set before saving the file as NavigationHyperlink.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel internal hyperlink | worksheet navigation link | Hyperlinks.Add | link to another sheet | Excel workbook navigation | Aspose.Cells example
// Common Searches: Aspose.Cells add hyperlink to another worksheet C# | How to create internal Excel links with Aspose.Cells | C# navigation hyperlink between sheets Aspose | Aspose.Cells Hyperlinks.Add usage example | Create table of contents with worksheet links in Aspose.Cells
// Developer Intent: Insert hyperlinks that jump to specific cells on other worksheets within the same Excel workbook using Aspose.Cells for .NET.
// Use Cases: Provide a quick link from Sheet1!A1 to Sheet2!A1 for easy navigation. | Add a back‑link on Sheet2!A1 to return to Sheet1!A1, enabling bidirectional movement. | Generate a workbook that serves as a clickable table of contents for multiple sheets. | Create interactive Excel reports where users can jump between summary and detail sheets.
// AI Prompts: Show me C# code that adds an internal hyperlink from one worksheet cell to another using Aspose.Cells. | Write an Aspose.Cells example that creates a workbook with two sheets and navigation links with custom display text. | Explain the parameters of Hyperlinks.Add for linking to a different sheet in an Excel file. | How can I build a table of contents with clickable sheet links in Aspose.Cells for .NET?

using System;
using Aspose.Cells;

namespace AsposeCellsNavigationHyperlink
{
    // This C# example creates a workbook, adds two sheets (Sheet1 and Sheet2), writes link text in cell A1 of each sheet, and uses the Hyperlinks.Add method to insert clickable links that jump to the opposite sheet's A1 cell. Custom display text is set before saving the file as NavigationHyperlink.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default first worksheet and rename it
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Put display text in Sheet1!A1
            sheet1.Cells["A1"].PutValue("Go to Sheet2");

            // Add a hyperlink in Sheet1!A1 that points to Sheet2!A1
            int linkIndex1 = sheet1.Hyperlinks.Add("A1", 1, 1, "Sheet2!A1");
            // Set custom text to display for the hyperlink
            sheet1.Hyperlinks[linkIndex1].TextToDisplay = "Go to Sheet2!A1";

            // Put display text in Sheet2!A1
            sheet2.Cells["A1"].PutValue("Back to Sheet1");

            // Add a hyperlink in Sheet2!A1 that points back to Sheet1!A1
            int linkIndex2 = sheet2.Hyperlinks.Add("A1", 1, 1, "Sheet1!A1");
            sheet2.Hyperlinks[linkIndex2].TextToDisplay = "Back to Sheet1!A1";

            // Save the workbook
            workbook.Save("NavigationHyperlink.xlsx");
        }
    }
}
