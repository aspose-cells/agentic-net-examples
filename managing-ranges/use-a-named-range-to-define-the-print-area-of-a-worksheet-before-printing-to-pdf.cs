using Aspose.Cells;
using System;

class PrintAreaFromNamedRange
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data
        worksheet.Cells["A1"].PutValue("Header1");
        worksheet.Cells["B1"].PutValue("Header2");
        worksheet.Cells["A2"].PutValue("Item1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Item2");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("Item3");
        worksheet.Cells["B4"].PutValue(300);

        // Create a named range that covers the area to be printed
        int nameIndex = workbook.Worksheets.Names.Add("PrintRange");
        Name namedRange = workbook.Worksheets.Names[nameIndex];
        // RefersTo must include the sheet name and start with '='
        namedRange.RefersTo = $"={worksheet.Name}!A1:B4";

        // Set the worksheet's print area using the address of the named range
        // GetRange() returns the actual range object; its Address property gives "A1:B4"
        string printAreaAddress = namedRange.GetRange().Address;
        worksheet.PageSetup.PrintArea = printAreaAddress;

        // Save the workbook as PDF; the defined print area will be used
        workbook.Save("PrintAreaFromNamedRange.pdf", SaveFormat.Pdf);
    }
}