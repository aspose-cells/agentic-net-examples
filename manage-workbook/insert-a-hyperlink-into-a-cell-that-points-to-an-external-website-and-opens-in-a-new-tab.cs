// Title: Add an external hyperlink to cell B2 that opens in a new browser tab using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add hyperlink to a specific cell that opens in a new tab | How to set a screen tip for an Excel hyperlink with Aspose.Cells | Saving an Excel workbook that contains external links using Aspose.Cells for .NET | Aspose.Cells Hyperlinks.Add method usage example in C# | Create a clickable URL in an Excel cell programmatically with Aspose.Cells | Open an external website from an Excel cell in a new browser window using Aspose.Cells
// Tags: Aspose.Cells create cell hyperlink C# | Excel hyperlink new tab Aspose.Cells | Aspose.Cells set hyperlink screen tip | Hyperlinks.Add API Aspose.Cells | Save workbook with external link Aspose.Cells | programmatic Excel URL insertion C#

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to create a new workbook, access the first worksheet, and insert an external hyperlink into cell B2 using Aspose.Cells for .NET. It sets the display text, adds a screen tip, and saves the workbook as HyperlinkDemo.xlsx, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Specify the cell (B2) where the hyperlink will be placed
            int row = 1;      // zero‑based index for row 2
            int column = 1;   // zero‑based index for column B

            // Add a hyperlink that points to an external website.
            // The Add method returns the index of the created hyperlink.
            int hyperlinkIndex = sheet.Hyperlinks.Add("https://www.example.com", row, column, "Visit Example");

            // Retrieve the Hyperlink object using the returned index
            Hyperlink link = sheet.Hyperlinks[hyperlinkIndex];

            // Optional: add a screen tip
            link.ScreenTip = "Click to open Example.com";

            // Define output file path
            string outputPath = "HyperlinkDemo.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
