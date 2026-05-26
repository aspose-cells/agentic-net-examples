using System;
using System.IO;
using Aspose.Cells;

class ClearNamedRangeContents
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ReportData"
            Name reportName = workbook.Worksheets.Names["ReportData"];
            if (reportName == null)
            {
                Console.WriteLine("Named range 'ReportData' not found.");
                return;
            }

            // Get the Aspose.Cells.Range object that the name refers to
            Aspose.Cells.Range reportRange = reportName.GetRange();

            // Clear only the contents, preserving formatting and the range definition
            reportRange.ClearContents();

            // Save the workbook with cleared contents
            workbook.Save(outputPath);

            Console.WriteLine($"Contents of named range 'ReportData' have been cleared and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}