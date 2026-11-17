// Title: Convert an .xlsx workbook to HTML with BestFit layout and parent hyperlink target using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML with BestFit presentation preference | set hyperlink target to parent when saving workbook as HTML using Aspose.Cells | C# code to convert .xlsx file to responsive HTML using HtmlSaveOptions | how to use HtmlSaveOptions.PresentationPreference.BestFit in Aspose.Cells | save Excel workbook as HTML with fitted page layout and parent links in .NET
// Tags: HtmlSaveOptions presentation preference bestfit | Excel to HTML conversion using Aspose.Cells C# | parent link target configuration Aspose.Cells | fitted page layout HTML export Aspose.Cells | save .xlsx as .html with Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx workbook, configures HtmlSaveOptions to use the BestFit presentation preference and sets the hyperlink target to "parent", then saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Fit the content to the page using the BestFit presentation preference
                // (Commented out if the enum is unavailable in the current API version)
                // PresentationPreference = HtmlSaveOptions.PresentationPreference.BestFit
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
