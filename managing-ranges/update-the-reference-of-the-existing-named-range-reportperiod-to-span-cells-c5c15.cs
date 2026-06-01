using Aspose.Cells;
using System;
using System.IO;

class UpdateNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "ReportPeriod";

            // Verify input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range; ensure it exists
            Name reportPeriod = workbook.Worksheets.Names[rangeName];
            if (reportPeriod == null)
                throw new InvalidOperationException($"Named range \"{rangeName}\" does not exist in the workbook.");

            // Determine the worksheet the name belongs to (0 = global, otherwise sheet index)
            string sheetName = workbook.Worksheets[reportPeriod.SheetIndex].Name;

            // Update the reference to span cells C5:C15 on that worksheet
            reportPeriod.RefersTo = $"={sheetName}!$C$5:$C$15";

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}