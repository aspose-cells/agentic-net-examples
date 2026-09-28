// Title: Update the 'ReportPeriod' named range to reference cells C5:C15 using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, finds the named range 'ReportPeriod', and changes its RefersTo address to C5:C15 on the same worksheet. | Show a .NET example that verifies the input .xlsx file exists, updates an existing named range's address, and saves the modified workbook to a new file. | Demonstrate how to parse the sheet name from a named range's RefersTo string and programmatically assign a new range reference using Aspose.Cells.
// Common Searches: Aspose.Cells C# update RefersTo for named range in .xlsx | Programmatically set named range to C5:C15 with Aspose.Cells | C# example to modify existing named range sheet reference using Aspose | Changing Excel named range address using Aspose.Cells .NET library
// Tags: Aspose.Cells modify named range address | C# update RefersTo property Aspose | Aspose.Cells set named range C5:C15 | Excel named range manipulation Aspose.Cells | Aspose.Cells workbook save after range change

using System;
using System.IO;
using Aspose.Cells;

// The sample loads input.xlsx, locates the named range 'ReportPeriod', extracts its worksheet name from the current RefersTo string, updates the range to C5:C15 on that sheet, and saves the workbook as output.xlsx while handling missing files and exceptions.
class UpdateNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Access the named range "ReportPeriod"
            var namedRange = workbook.Worksheets.Names["ReportPeriod"];
            if (namedRange != null)
            {
                // Extract the original worksheet name from the existing RefersTo string
                // RefersTo format: SheetName!A1:B2
                string originalRef = namedRange.RefersTo ?? string.Empty;
                string sheetName = originalRef.Split('!')[0];

                // Update the range to refer to cells C5:C15 on the same worksheet
                namedRange.RefersTo = $"{sheetName}!C5:C15";
            }
            else
            {
                Console.WriteLine("Named range 'ReportPeriod' not found.");
            }

            // Save the workbook with the updated named range
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
