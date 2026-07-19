// Title: C# Aspose.Cells Example: Create a Workbook, Add a "SalesData" Sheet, Fill Cells, and Save as XLSX
// Description: Demonstrates how to instantiate an Aspose.Cells Workbook in C#, add a worksheet named "SalesData", write header labels and three rows of monthly sales records (Month, Product, Units Sold, Revenue) into cells A1:D4, and export the file as SalesData.xlsx.
// Keywords: Aspose.Cells C# workbook creation | add worksheet Aspose.Cells | populate Excel cells with data | save workbook as XLSX .NET | sales data Excel example | Aspose.Cells tutorial | C# Excel automation | global developers Aspose.Cells
// Common Searches: How to create an Excel file with Aspose.Cells in C# | Aspose.Cells add worksheet and write data example | C# code to save sales data to XLSX using Aspose.Cells | Aspose.Cells sample for populating cells programmatically | Export monthly sales records to Excel with Aspose.Cells
// Developer Intent: Generate an Excel workbook that contains a pre‑filled "SalesData" worksheet for quick reporting or further processing.
// Use Cases: Automated generation of monthly sales reports for business intelligence. | Creating a reusable Excel template with predefined headers for data entry workflows. | Exporting sales records from a database or API into a portable XLSX file for stakeholder review.
// AI Prompts: Show how to apply bold font and background color to the header row using Aspose.Cells. | Provide a snippet that adds a column chart based on the Units Sold and Revenue columns. | Explain how to loop through a collection of sales objects and write each item to the worksheet dynamically.

using System;
using Aspose.Cells;

namespace SalesDataWorkbook
{
    // Demonstrates how to instantiate an Aspose.Cells Workbook in C#, add a worksheet named "SalesData", write header labels and three rows of monthly sales records (Month, Product, Units Sold, Revenue) into cells A1:D4, and export the file as SalesData.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Add a new worksheet named "SalesData"
            Worksheet sheet = workbook.Worksheets.Add("SalesData");

            // Populate header row
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Product");
            sheet.Cells["C1"].PutValue("Units Sold");
            sheet.Cells["D1"].PutValue("Revenue");

            // Sample sales data
            sheet.Cells["A2"].PutValue("January");
            sheet.Cells["B2"].PutValue("Widget");
            sheet.Cells["C2"].PutValue(120);
            sheet.Cells["D2"].PutValue(2400);

            sheet.Cells["A3"].PutValue("February");
            sheet.Cells["B3"].PutValue("Gadget");
            sheet.Cells["C3"].PutValue(85);
            sheet.Cells["D3"].PutValue(2550);

            sheet.Cells["A4"].PutValue("March");
            sheet.Cells["B4"].PutValue("Doohickey");
            sheet.Cells["C4"].PutValue(150);
            sheet.Cells["D4"].PutValue(3750);

            // Save the workbook to a file
            workbook.Save("SalesData.xlsx", SaveFormat.Xlsx);
        }
    }
}
