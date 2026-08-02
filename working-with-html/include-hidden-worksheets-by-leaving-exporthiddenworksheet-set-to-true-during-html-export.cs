// Title: Export Hidden Worksheets to HTML using Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook with visible and hidden sheets, enable HtmlSaveOptions.ExportHiddenWorksheet, and save the entire workbook as an HTML file that includes hidden worksheets.
// Keywords: Aspose.Cells HTML export hidden worksheet | ExportHiddenWorksheet C# | save workbook as HTML .NET | include hidden sheets Aspose.Cells | HtmlSaveOptions ExportHiddenWorksheet true | Aspose.Cells export entire workbook to HTML | C# Aspose.Cells hidden worksheet export
// Common Searches: Aspose.Cells export hidden worksheets to HTML | HtmlSaveOptions ExportHiddenWorksheet property C# | How to include hidden sheets when saving as HTML with Aspose.Cells | Export entire workbook including hidden sheets Aspose.Cells .NET | C# save workbook as HTML with hidden worksheets
// Developer Intent: The developer wants to generate an HTML file that contains both visible and hidden worksheets from an Aspose.Cells workbook, preserving hidden data in the output.
// Use Cases: Web dashboard that displays analysis data stored in hidden sheets alongside visible data. | Archiving Excel workbooks as HTML while retaining hidden calculation worksheets for future reference. | Creating printable HTML reports that include every worksheet, even those hidden in the original file. | Sharing workbook content with users who lack Excel but need to view hidden data in a browser.
// AI Prompts: Provide C# code to export only selected hidden worksheets to HTML with custom CSS using Aspose.Cells. | Explain performance considerations when exporting large workbooks that contain hidden sheets to HTML. | Show how to combine ExportHiddenWorksheet with other HtmlSaveOptions such as EmbedImages and PageSetup. | Give steps to programmatically toggle worksheet visibility before HTML export with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook with visible and hidden sheets, enable HtmlSaveOptions.ExportHiddenWorksheet, and save the entire workbook as an HTML file that includes hidden worksheets.
    public class ExportHiddenWorksheetsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Visible worksheet with sample data
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Visible Data");

                // Add a hidden worksheet and put data into it
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("Hidden Data");
                hiddenSheet.IsVisible = false; // hide the worksheet

                // Set HTML save options to export hidden worksheets
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true,      // include hidden sheets in the output
                    ExportActiveWorksheetOnly = false // export the entire workbook
                };

                // Determine output file path
                string outputPath = "output_with_hidden.html";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML with hidden worksheets included
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting hidden worksheets:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportHiddenWorksheetsDemo.Run();
        }
    }
}
