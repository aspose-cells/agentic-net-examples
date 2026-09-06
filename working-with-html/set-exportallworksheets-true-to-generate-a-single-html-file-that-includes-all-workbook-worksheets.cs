// Title: Export every worksheet of an Excel workbook to a single HTML file using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Create C# code that configures Aspose.Cells HtmlSaveOptions to generate one HTML document that contains all worksheets of a workbook. | Update the sample to turn on the option that includes every sheet when saving the workbook as HTML with Aspose.Cells. | Demonstrate how to save a .NET Workbook as a combined HTML page covering all sheets using Aspose.Cells.
// Common Searches: Aspose.Cells C# generate single HTML file from all workbook sheets | How to include every worksheet when converting Excel to HTML with Aspose.Cells | HtmlSaveOptions setting to export all sheets in .NET | Combine multiple Excel worksheets into one HTML page using Aspose.Cells library | Save Excel workbook as HTML with all worksheets in C#
// Tags: Aspose.Cells HtmlSaveOptions ExportAllWorksheets | C# save workbook as single HTML | include all worksheets in HTML conversion | Aspose.Cells multiple sheets HTML output | HTML export option for entire Excel workbook

using System;
using Aspose.Cells;

// The example creates a workbook with two worksheets, fills cells with sample data, sets HtmlSaveOptions to include every worksheet, and saves the workbook as a single HTML file named AllWorksheets.html using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add data to the first worksheet
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Hello");
            sheet1.Cells["B1"].PutValue("World");

            // Add a second worksheet with some data
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheetIndex];
            sheet2.Name = "Sheet2";
            sheet2.Cells["A1"].PutValue(123);
            sheet2.Cells["B1"].PutValue(456);

            // Configure HTML save options to export all worksheets into a single HTML file
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Export all worksheets (set to false to include all)
            saveOptions.ExportActiveWorksheetOnly = false;

            // Save the workbook as HTML
            string outputPath = "AllWorksheets.html";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
