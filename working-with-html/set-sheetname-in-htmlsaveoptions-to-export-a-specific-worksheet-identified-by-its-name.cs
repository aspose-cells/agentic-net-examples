// Title: Export a single worksheet to HTML by specifying its name with HtmlSaveOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export only one worksheet to HTML by sheet name | How to set HtmlSaveOptions SheetName to save a specific worksheet as HTML | C# export selected worksheet to HTML using Aspose.Cells | HtmlSaveOptions ExportActiveWorksheetOnly vs SheetName property in Aspose.Cells | Save particular worksheet as HTML with Aspose.Cells .NET API
// Tags: HtmlSaveOptions sheet name export | export single worksheet to HTML Aspose.Cells | Aspose.Cells C# HTML export selected sheet | set active worksheet only HtmlSaveOptions | Aspose.Cells HTML export options C#

using System;
using Aspose.Cells;

// The example creates a workbook with two sheets, activates the "Details" sheet, configures HtmlSaveOptions (ExportActiveWorksheetOnly = true, ExportGridLines = true, ExportImagesAsBase64 = true) and saves that specific worksheet as an HTML file.
class Program
{
    static void Main()
    {
        const string outputHtml = "ExportedSheet.html";

        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();

            // First worksheet (default name: Sheet1)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";
            sheet1.Cells["A1"].PutValue("Product");
            sheet1.Cells["B1"].PutValue("Quantity");
            sheet1.Cells["A2"].PutValue("Apple");
            sheet1.Cells["B2"].PutValue(150);
            sheet1.Cells["A3"].PutValue("Banana");
            sheet1.Cells["B3"].PutValue(200);

            // Second worksheet (the one to export)
            Worksheet sheet2 = workbook.Worksheets.Add("Details");
            sheet2.Cells["A1"].PutValue("Order ID");
            sheet2.Cells["B1"].PutValue("Customer");
            sheet2.Cells["C1"].PutValue("Amount");
            sheet2.Cells["A2"].PutValue(1001);
            sheet2.Cells["B2"].PutValue("John Doe");
            sheet2.Cells["C2"].PutValue(123.45);
            sheet2.Cells["A3"].PutValue(1002);
            sheet2.Cells["B3"].PutValue("Jane Smith");
            sheet2.Cells["C3"].PutValue(678.90);

            // Make the desired worksheet the active one
            workbook.Worksheets.ActiveSheetIndex = sheet2.Index;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,
                ExportImagesAsBase64 = true,
                ExportActiveWorksheetOnly = true
            };

            // Save the selected worksheet as an HTML file
            workbook.Save(outputHtml, htmlOptions);
            Console.WriteLine($"Worksheet \"{sheet2.Name}\" has been exported to \"{outputHtml}\" successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while exporting the worksheet:");
            Console.WriteLine(ex.Message);
        }
    }
}
