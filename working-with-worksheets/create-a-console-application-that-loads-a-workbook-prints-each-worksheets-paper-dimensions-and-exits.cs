using System;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Path to the workbook; can be passed as a command‑line argument
        string filePath = args.Length > 0 ? args[0] : "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Iterate through all worksheets and output their paper dimensions
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            PageSetup setup = sheet.PageSetup;

            Console.WriteLine($"Worksheet {i} - \"{sheet.Name}\": Width = {setup.PaperWidth} inches, Height = {setup.PaperHeight} inches");
        }
    }
}