// Title: Define a named range as the print area and save a worksheet as PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set print area from a named range before exporting to PDF | How to assign Worksheet.PageSetup.PrintArea using a workbook name collection in .NET | Create a named range and use it as the print area for PDF conversion with Aspose.Cells | Fit printed content to a single page when saving Excel to PDF in Aspose.Cells | Programmatically set print area in Aspose.Cells using Name.RefersTo
// Tags: named range print area Aspose.Cells | worksheet PageSetup.PrintArea C# | export worksheet to PDF Aspose.Cells | fit to one page PDF Aspose.Cells | add workbook name collection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, populates cells A1:C10, defines a named range called "PrintAreaRange", assigns that range to the worksheet's PageSetup.PrintArea, configures the page to fit on one sheet, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some sample data (A1:C10)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue(i - 1);
                sheet.Cells[$"B{i}"].PutValue((i - 1) * 10);
                sheet.Cells[$"C{i}"].PutValue((i - 1) * 100);
            }

            // Define a named range that will be used as the print area
            // The reference must include the sheet name and absolute cell addresses
            string rangeReference = $"{sheet.Name}!$A$1:$C$10";

            // Add the name to the workbook's name collection (Add overload takes only the name)
            int namedRangeIndex = workbook.Worksheets.Names.Add("PrintAreaRange");
            Name printAreaRange = workbook.Worksheets.Names[namedRangeIndex];
            printAreaRange.RefersTo = rangeReference;

            // Set the worksheet's print area to the named range reference
            sheet.PageSetup.PrintArea = printAreaRange.RefersTo;

            // Optional: fit the printed content to a single page
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 1;

            // Determine output file path and ensure the directory exists
            string outputPath = "Output.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
