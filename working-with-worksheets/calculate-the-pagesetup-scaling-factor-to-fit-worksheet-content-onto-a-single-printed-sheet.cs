// Title: How to retrieve the page‑setup Zoom scaling factor after fitting an Aspose.Cells worksheet to a single printed page in C#
// AI Prompts: Create C# code that sets Worksheet.PageSetup.FitToPagesWide = 1 and FitToPagesTall = 1 with Aspose.Cells, then reads the Worksheet.PageSetup.Zoom value and prints it. | Write a C# example that configures a workbook to print on one page using Aspose.Cells and returns the calculated scaling percentage from the Zoom property.
// Common Searches: Aspose.Cells .NET get Zoom percentage after setting FitToPagesWide and FitToPagesTall | C# determine print scaling factor for Excel worksheet using Aspose.Cells page setup | How to read page setup scaling factor when fitting worksheet to one page in Aspose.Cells | Retrieve worksheet zoom value after FitToPagesWide=1 in Aspose.Cells C#
// Tags: Aspose.Cells page setup FitToPagesWide property | Aspose.Cells page setup FitToPagesTall usage | Aspose.Cells retrieve Zoom scaling | C# Aspose.Cells print scaling factor | Aspose.Cells fit worksheet on one printed page

using Aspose.Cells;
using System;

// Demonstrates creating a workbook, populating data, configuring the worksheet to fit on a single printed page with FitToPagesWide and FitToPagesTall, reading the resulting Zoom scaling factor, and outputting the percentage.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(200);
            // Add more rows as needed...

            // Configure page setup to fit the entire sheet on a single printed page
            sheet.PageSetup.FitToPagesWide = 1; // fit width to 1 page
            sheet.PageSetup.FitToPagesTall = 1; // fit height to 1 page

            // After setting FitToPages, the Zoom property reflects the scaling factor (percentage)
            double scalingFactor = sheet.PageSetup.Zoom;

            // Output the scaling factor
            Console.WriteLine("Scaling factor to fit on one page: " + scalingFactor + "%");

            // Save the workbook (optional)
            workbook.Save("ScaledWorksheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
