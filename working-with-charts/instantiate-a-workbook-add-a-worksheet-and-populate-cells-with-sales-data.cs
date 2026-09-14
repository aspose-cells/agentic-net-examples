// Title: Create an Excel workbook with a 'SalesData' worksheet and fill cells A1:C4 with product, region, and sales values using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to instantiate a Workbook, add a worksheet named 'SalesData', write a header row and three rows of product‑region‑sales data into cells A1:C4, and save the file as an XLSX. | Show how to programmatically set cell values for product, region, and sales columns in a new Excel file with Aspose.Cells in a .NET console application. | Demonstrate adding a worksheet, populating it with sample sales records, and exporting the workbook to 'SalesData.xlsx' using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# create workbook and add worksheet named SalesData | populate Excel cells A1 to C4 with sample sales data using Aspose.Cells .NET | how to write header row and data rows in Excel with Aspose.Cells C# | save a new Excel file as XLSX with Aspose.Cells in a console app | C# example for inserting product, region, sales values into an Excel sheet using Aspose.Cells
// Tags: create workbook Aspose.Cells C# | add worksheet and populate cells Aspose.Cells | write header and data rows Excel Aspose.Cells | save workbook as XLSX Aspose.Cells | sample sales dataset Excel generation C#

using System;
using Aspose.Cells;

// // Creates a new Excel workbook, adds a 'SalesData' worksheet, writes a header and three rows of product/region/sales data into cells A1:C4, and saves the file as SalesData.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the default worksheet and give it a meaningful name
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SalesData";

        // Add header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Sales");

        // Populate sample sales data
        sheet.Cells["A2"].PutValue("Laptop");
        sheet.Cells["B2"].PutValue("North");
        sheet.Cells["C2"].PutValue(1200);

        sheet.Cells["A3"].PutValue("Smartphone");
        sheet.Cells["B3"].PutValue("South");
        sheet.Cells["C3"].PutValue(850);

        sheet.Cells["A4"].PutValue("Tablet");
        sheet.Cells["B4"].PutValue("East");
        sheet.Cells["C4"].PutValue(430);

        // Save the workbook to an XLSX file
        workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
    }
}
