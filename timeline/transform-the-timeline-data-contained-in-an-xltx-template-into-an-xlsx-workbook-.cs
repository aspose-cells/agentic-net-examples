using System;
using Aspose.Cells;
using Aspose.Cells.Timelines;

class TransformTemplateToWorkbook
{
    static void Main()
    {
        // Path to the XLTX template that contains the Timeline control
        string templatePath = "Template.xltx";

        // Load the template workbook. The constructor automatically detects the file format.
        Workbook workbook = new Workbook(templatePath);

        // At this point the workbook contains all original content,
        // including the Timeline objects that were defined in the template.

        // Optionally, you can access the Timeline collection to verify it was loaded.
        // Example: print the number of timelines in the first worksheet.
        Worksheet firstSheet = workbook.Worksheets[0];
        Console.WriteLine($"Timelines loaded: {firstSheet.Timelines.Count}");

        // Save the workbook as a regular XLSX file, preserving all content.
        string outputPath = "Result.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}