// Title: Set a worksheet's print area to a named range before converting an Excel workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set print area from a named range before PDF export | How to limit PDF output to a specific named range using Aspose.Cells | C# convert Excel to PDF with only the PrintRange area using Aspose.Cells | Programmatically assign worksheet PageSetup.PrintArea to a named range in Aspose.Cells | Save Excel workbook as PDF with custom print area defined by a name in .NET
// Tags: Aspose.Cells set named range as print area | C# PDF export limited to named range | PageSetup.PrintArea configuration Aspose.Cells | Excel to PDF conversion using named range print area | Aspose.Cells workbook save PDF print area

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, retrieves the named range "PrintRange", assigns its address to the worksheet's PageSetup.PrintArea, and then saves the workbook as a PDF so that only the defined range appears in the output.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the named range definition from the workbook's Names collection
            Name printRangeName = workbook.Worksheets.Names["PrintRange"];
            if (printRangeName == null)
            {
                Console.WriteLine("Error: Named range \"PrintRange\" not found.");
                return;
            }

            // Get the address of the named range (e.g., "=Sheet1!$A$1:$D$10")
            string rangeRef = printRangeName.RefersTo;

            // Remove the leading '=' if present
            if (!string.IsNullOrEmpty(rangeRef) && rangeRef.StartsWith("="))
                rangeRef = rangeRef.Substring(1);

            // Set the print area to the named range address
            sheet.PageSetup.PrintArea = rangeRef;

            // Save the workbook as PDF – only the defined print area will be included
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
