// Title: Create an internal hyperlink that jumps to Sheet1!A1 in an Aspose.Cells workbook using C#
// AI Prompts: Insert a hyperlink into cell B2 that navigates to '#Sheet1!A1' and set its displayed text to 'Go to Sheet1!A1' with Aspose.Cells for .NET. | Generate a new workbook, add an internal link pointing to a different worksheet cell range, and save it as HyperlinkInternal.xlsx using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# add internal hyperlink to another worksheet cell | How to use '#Sheet1!A1' syntax for internal links in Aspose.Cells | C# example of creating a hyperlink that opens a specific sheet and cell in an Excel file with Aspose
// Tags: Aspose.Cells internal hyperlink creation | C# hyperlink to worksheet cell reference | Aspose.Cells '#Sheet1!A1' link syntax | programmatic Excel hyperlink .NET

using System;
using Aspose.Cells;

// The program creates a new workbook, adds an internal hyperlink in cell B2 that points to Sheet1!A1 using the '#Sheet1!A1' reference, sets custom display text, and saves the file as HyperlinkInternal.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell B2 that points to Sheet1!A1 using the '#Sheet1!A1' syntax
        int hyperlinkIndex = worksheet.Hyperlinks.Add("B2", 1, 1, "#Sheet1!A1");

        // Set the text that will be displayed for the hyperlink
        worksheet.Hyperlinks[hyperlinkIndex].TextToDisplay = "Go to Sheet1!A1";

        // Save the workbook
        workbook.Save("HyperlinkInternal.xlsx");
    }
}
