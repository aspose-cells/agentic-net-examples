// Title: Define Print Area with a Named Range and Export to PDF using Aspose.Cells for .NET
// Description: Creates a workbook, adds a named range (MyPrintArea) covering cells A1:B3, assigns the range address to PageSetup.PrintArea, and saves the worksheet as a PDF so only the named range is printed.
// Keywords: Aspose.Cells print area named range | C# set print area Aspose.Cells | export worksheet range to PDF C# | Aspose.Cells PageSetup.PrintArea | named range PDF export Aspose.Cells | Aspose.Cells .NET PDF generation
// Common Searches: how to set a named range as print area Aspose.Cells | Aspose.Cells export selected range to PDF C# | C# Aspose.Cells print area from named range example | save only part of a worksheet as PDF using Aspose.Cells | Aspose.Cells PageSetup.PrintArea with named range
// Developer Intent: Set a named range as the worksheet's print area and generate a PDF containing only that range.
// Use Cases: Produce a printable report that includes just the header and data rows. | Create individual invoice PDFs by defining each invoice section as a named range. | Automate a summary sheet that prints only a specific table without extra worksheet content.
// AI Prompts: Generate C# code that builds a dynamic named range based on the last used row and uses it as the print area. | Show how to assign different named ranges as print areas for multiple worksheets and export each to a separate PDF. | Explain how to retrieve a named range address at runtime when the range name is supplied by the user and set it as the print area.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsPrintAreaDemo
{
    // Creates a workbook, adds a named range (MyPrintArea) covering cells A1:B3, assigns the range address to PageSetup.PrintArea, and saves the worksheet as a PDF so only the named range is printed.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that will be part of the print area
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Data1");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["A3"].PutValue("Data2");
                sheet.Cells["B3"].PutValue(200);

                // -----------------------------------------------------------------
                // Define a named range that covers the cells we want to print
                // -----------------------------------------------------------------
                // Add a new name to the workbook's name collection
                int nameIndex = workbook.Worksheets.Names.Add("MyPrintArea");
                // Set the reference of the named range (including the sheet name)
                // Note: the reference must start with '=' as per Excel formula syntax
                workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$B$3";

                // Retrieve the Range object represented by the named range
                Name namedRange = workbook.Worksheets.Names["MyPrintArea"];
                AsposeRange range = namedRange.GetRange();

                // The PrintArea property expects a plain address without the leading '='
                // Range.Address returns the address in A1 style (e.g., "A1:B3")
                sheet.PageSetup.PrintArea = range.Address;

                // -----------------------------------------------------------------
                // Save the workbook to PDF – only the defined print area will be exported
                // -----------------------------------------------------------------
                string outputPath = "PrintAreaOutput.pdf";
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF generated with named range as print area: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
