// Title: How to set an Excel worksheet’s page orientation to landscape using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet print orientation to landscape | Change page setup orientation to landscape in an Excel file with Aspose.Cells | C# code to make Excel worksheet print in landscape mode using Aspose.Cells | Set landscape page orientation for a specific worksheet in Aspose.Cells .NET | Programmatically set Excel worksheet orientation to landscape with Aspose.Cells API
// Tags: Aspose.Cells worksheet page setup orientation | C# configure Excel print layout landscape | Aspose.Cells API set page orientation | Save workbook with landscape orientation .NET | Excel worksheet landscape printing Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates creating a workbook, adding sample data, setting the worksheet’s page setup orientation to landscape, and saving the file as an .xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or you could load an existing file)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Add some sample data (optional, just to illustrate)
            worksheet.Cells["A1"].PutValue("Item");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apples");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Oranges");
            worksheet.Cells["B3"].PutValue(15);

            // Set the page orientation of this worksheet to Landscape
            worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Save the workbook to a file
            string outputFile = "LandscapeOrientation.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputFile}' with landscape orientation.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
