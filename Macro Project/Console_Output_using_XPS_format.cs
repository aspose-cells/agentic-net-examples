using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data to the worksheet
        sheet.Cells["A1"].PutValue("Console Output using XPS format");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Create XpsSaveOptions using the provided constructor
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        // Configure desired options (optional)
        saveOptions.OnePagePerSheet = true;   // each sheet on a single XPS page
        saveOptions.DefaultFont = "Arial";   // fallback font for Unicode characters

        // Define output file name
        string outputFile = "ConsoleOutputDemo.xps";

        // Save the workbook as XPS using the options
        workbook.Save(outputFile, saveOptions);

        // Write confirmation to console
        Console.WriteLine($"Workbook successfully saved as XPS to: {outputFile}");
    }
}