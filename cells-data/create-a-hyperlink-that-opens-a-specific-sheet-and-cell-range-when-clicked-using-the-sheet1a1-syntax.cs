// Title: Aspose.Cells for .NET – Add an internal hyperlink to a worksheet cell using #Sheet1!A1 syntax
// Description: C# example that creates a new workbook, adds a second worksheet, writes a value to A1, and inserts an internal hyperlink in B2 of the first sheet that points to "#Sheet2!A1". The hyperlink text is set to "Go to Sheet2 A1" and the file is saved as HyperlinkInternal.xlsx.
// Keywords: Aspose.Cells | C# | internal hyperlink | Excel navigation | #SheetName!A1 | hyperlink text | Hyperlinks.Add | worksheet link | Aspose.Cells API | Excel internal link
// Common Searches: Aspose.Cells add internal hyperlink C# | Excel #Sheet1!A1 hyperlink syntax | How to link to another worksheet cell with Aspose.Cells | Set hyperlink display text in Aspose.Cells workbook | Create navigation links in Excel using Aspose.Cells
// Developer Intent: Insert an internal hyperlink that opens a specific worksheet and cell when the user clicks it.
// Use Cases: Build a table‑of‑contents sheet that jumps to key cells on other worksheets. | Create a dashboard with quick links to result cells across multiple sheets. | Enable readers to navigate from a summary page to detailed data sections instantly.
// AI Prompts: Generate C# code with Aspose.Cells to add a hyperlink from Sheet1!C5 to Sheet3!D10 and set custom display text. | Explain how to modify the target sheet and cell of an existing Aspose.Cells hyperlink at runtime based on user input. | Show a loop that reads a list of sheet names and creates internal hyperlinks on a navigation sheet using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// C# example that creates a new workbook, adds a second worksheet, writes a value to A1, and inserts an internal hyperlink in B2 of the first sheet that points to "#Sheet2!A1". The hyperlink text is set to "Go to Sheet2 A1" and the file is saved as HyperlinkInternal.xlsx.
class HyperlinkToSheetCell
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (source sheet where the hyperlink will be placed)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a second worksheet that will be the hyperlink target.
            // Use a unique name to avoid conflict with the default sheet name.
            string targetSheetName = "Sheet2";
            Worksheet targetSheet = workbook.Worksheets.Add(targetSheetName);
            targetSheet.Cells["A1"].PutValue("Target Cell");

            // Add a hyperlink in cell B2 of the source sheet that points to Sheet2!A1
            // The address uses the '#SheetName!Cell' syntax for an internal cell reference
            int hyperlinkIndex = sourceSheet.Hyperlinks.Add("B2", 1, 1, $"#{targetSheetName}!A1");

            // Set the text that will be displayed for the hyperlink
            sourceSheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Go to Sheet2 A1";

            // Define output file path
            string outputPath = "HyperlinkInternal.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
