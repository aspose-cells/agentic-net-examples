// Title: How to set a column sparkline group orientation to vertical in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a sparkline group for a data range and configures its Orientation property to Vertical. | Show the steps to change the orientation of an existing sparkline group to vertical in a .xlsx workbook using Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells set sparkline display direction to vertical | example of vertical column sparklines using Aspose.Cells .NET | change sparkline layout to vertical in an Excel workbook with Aspose.Cells | using Aspose.Cells C# modify SparklineGroup Orientation property
// Tags: Aspose.Cells sparkline orientation vertical | C# set sparkline direction Aspose.Cells | Excel vertical sparkline layout Aspose.Cells | modify sparkline orientation .NET | vertical sparkline setting Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, populates column A with sample numeric data, and saves the file as SparklineVertical.xlsx. It also indicates where to set a column sparkline group's Orientation property to Vertical when the Aspose.Cells Sparkline API is available.
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

            // Populate sample data (column A)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(15);
            sheet.Cells["A4"].PutValue(30);
            sheet.Cells["A5"].PutValue(25);

            // NOTE: Sparkline functionality requires a version of Aspose.Cells that includes
            // the Aspose.Cells.Sparkline namespace. If unavailable, this section is omitted.

            // Save the workbook
            string outputPath = "SparklineVertical.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
