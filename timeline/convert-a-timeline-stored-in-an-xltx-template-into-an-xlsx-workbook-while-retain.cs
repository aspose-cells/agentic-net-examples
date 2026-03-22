using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Determine the full path to the template file located in the same directory as the executable.
        string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TimelineTemplate.xltx");

        if (!File.Exists(templatePath))
        {
            Console.WriteLine($"Template file not found: {templatePath}");
            return;
        }

        // Load the XLTX template that contains the Timeline.
        Workbook workbook = new Workbook(templatePath);

        // Save the workbook as an XLSX file.
        // All data, including the Timeline control, is preserved.
        string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TimelineOutput.xlsx");
        workbook.Save(outputPath);

        Console.WriteLine($"Workbook saved successfully to: {outputPath}");
    }
}