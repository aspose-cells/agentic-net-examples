// Title: Create a PDF from a specific named range in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export only a named range to PDF | How to set print area from a named range before saving Excel as PDF with Aspose.Cells | Generate PDF of a selected range in Excel using Aspose.Cells .NET | Aspose.Cells .NET save workbook as PDF limited to MyRange | Retrieve named range and convert to PDF with Aspose.Cells C#
// Tags: Aspose.Cells named range PDF conversion | C# GetRangeByName method Aspose.Cells | assign worksheet print area programmatically | Aspose.Cells SaveFormat.Pdf with selected range | export Excel worksheet to PDF limited area C#

using Aspose.Cells;
using System;
using System.IO;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel file, locates the named range “MyRange”, sets the worksheet’s print area to that range, activates the sheet, and saves the workbook as a PDF so only the named range appears in the output, with error handling for missing files or ranges.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range (replace "MyRange" with your actual range name)
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range \"MyRange\" not found.");
                return;
            }

            // Get the worksheet that contains the named range via the Range object
            AsposeRange range = workbook.Worksheets.GetRangeByName("MyRange");
            Worksheet sheet = range?.Worksheet ?? workbook.Worksheets[0];

            // Set the print area of the worksheet to the named range's reference
            sheet.PageSetup.PrintArea = namedRange.RefersTo;

            // Make this worksheet the active sheet so it is the one rendered to PDF
            workbook.Worksheets.ActiveSheetIndex = sheet.Index;

            // Save the workbook as a PDF; only the defined print area will be included
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
