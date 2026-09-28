// Title: Exclude hidden worksheets from HTML conversion with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that hides a worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet to false, saves the workbook as HTML, and checks that the hidden sheet name does not appear in the generated file. | Show how to programmatically verify that a hidden worksheet is omitted from the HTML output produced by Aspose.Cells by reading the file content and searching for the sheet name.
// Common Searches: Aspose.Cells C# export workbook to HTML without hidden sheets | How to prevent hidden worksheets from being saved in HTML using Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet false example in .NET | Verify hidden worksheet exclusion after HTML conversion with Aspose.Cells | C# code to hide worksheet and exclude it from HTML output Aspose
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# hide worksheet for HTML export | HTML conversion excluding hidden sheets | verify hidden worksheet omission Aspose.Cells | Aspose.Cells workbook to HTML without hidden worksheets

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, marks the first worksheet as hidden, configures HtmlSaveOptions.ExportHiddenWorksheet to false, saves the workbook as an HTML file, reads the generated HTML, and confirms that the hidden worksheet's name is not present, demonstrating how to exclude hidden sheets from HTML conversion using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure the first worksheet is hidden for the test
        Worksheet hiddenSheet = workbook.Worksheets[0];
        hiddenSheet.IsVisible = false;

        // Set HTML save options to exclude hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = false;

        // Save the workbook as HTML
        string htmlPath = "output.html";
        workbook.Save(htmlPath, htmlOptions);

        // Verify that the hidden worksheet name does not appear in the generated HTML
        string htmlContent = File.ReadAllText(htmlPath);
        bool hiddenSheetPresent = htmlContent.Contains(hiddenSheet.Name);
        Console.WriteLine(hiddenSheetPresent
            ? "Hidden worksheet was exported."
            : "Hidden worksheet correctly excluded from HTML.");
    }
}
