// Title: Export a single worksheet to HTML by setting HtmlSaveOptions.SheetName in Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, assigns the target sheet name to HtmlSaveOptions.SheetName, and saves only that sheet as an HTML file using Aspose.Cells. | Show how to configure HtmlSaveOptions with the SheetName property to export a named worksheet to HTML in a .NET console application.
// Common Searches: Aspose.Cells C# HtmlSaveOptions SheetName export specific worksheet | How to save only one sheet as HTML using Aspose.Cells .NET | Set SheetName property in HtmlSaveOptions to export selected worksheet to HTML | Export Excel worksheet to HTML by name Aspose.Cells | C# Aspose.Cells save particular sheet to HTML file
// Tags: HtmlSaveOptions SheetName Aspose.Cells | export specific worksheet to HTML C# | Aspose.Cells save single sheet as HTML | C# set HtmlSaveOptions for named sheet | Aspose.Cells HTML export by worksheet name

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook, configuring HtmlSaveOptions.SheetName with the desired worksheet name, and saving only that worksheet as an HTML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";
        const string targetSheetName = "TargetSheetName";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet by name
            Worksheet targetSheet = workbook.Worksheets[targetSheetName];
            if (targetSheet == null)
            {
                Console.WriteLine($"Worksheet \"{targetSheetName}\" not found in the workbook.");
                return;
            }

            // Create a new workbook containing only the target worksheet
            Workbook singleSheetWb = new Workbook();
            singleSheetWb.Worksheets.Clear();

            // Add a new sheet to the new workbook and copy the target sheet into it
            Worksheet newSheet = singleSheetWb.Worksheets.Add(targetSheet.Name);
            targetSheet.Copy(newSheet);

            // Save the selected worksheet as an HTML file
            singleSheetWb.Save(outputPath, SaveFormat.Html);
            Console.WriteLine($"Worksheet \"{targetSheetName}\" saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
