// Title: Export a specific worksheet by name to HTML with Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Generate C# code that loads an Excel file, selects a worksheet by its name, sets it as the active sheet, and saves only that sheet as an HTML file using Aspose.Cells HtmlSaveOptions. | Demonstrate how to configure HtmlSaveOptions.ExportActiveWorksheetOnly to true after assigning Workbook.Worksheets.ActiveSheetIndex to a named worksheet. | Include robust error handling for a missing input file and for a worksheet name that does not exist when exporting a single sheet to HTML.
// Common Searches: Aspose.Cells C# export only one worksheet to HTML by sheet name | How to use HtmlSaveOptions to save a specific sheet as HTML in .NET | Set active worksheet index for HtmlSaveOptions ExportActiveWorksheetOnly example | C# code to handle missing worksheet error when exporting to HTML with Aspose.Cells | Export selected worksheet to HTML using Aspose.Cells SaveFormat.Html
// Tags: htmlsaveoptions export selected worksheet aspocells | aspocells export named sheet to html c# | c# set active worksheet index aspocells | aspocells html export specific worksheet | error handling missing worksheet aspocells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source Excel file exists, loads it with Aspose.Cells, retrieves the worksheet named "Sheet2", makes it the active sheet, configures HtmlSaveOptions with ExportActiveWorksheetOnly=true, and saves only that worksheet as an HTML file while providing error handling for missing files and missing worksheets.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Locate the worksheet named "Sheet2"
            Worksheet sheet = workbook.Worksheets["Sheet2"];
            if (sheet == null)
            {
                Console.WriteLine("Error: Worksheet \"Sheet2\" does not exist in the workbook.");
                return;
            }

            // Set the active sheet to the target worksheet
            workbook.Worksheets.ActiveSheetIndex = sheet.Index;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true
            };

            // Save the selected worksheet as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Worksheet \"{sheet.Name}\" has been saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
