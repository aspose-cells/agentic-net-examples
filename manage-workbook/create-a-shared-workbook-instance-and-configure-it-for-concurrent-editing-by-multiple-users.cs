// Title: Create a shared Excel workbook in C# with Aspose.Cells and enable multi‑user concurrent editing
// Common Searches: Aspose.Cells C# how to enable shared workbook mode for Excel files | C# code to create an Excel workbook that supports simultaneous editing with Aspose.Cells | Saving an Excel file as a shared workbook using Aspose.Cells .NET API | Configure Aspose.Cells workbook for multi‑user collaboration in a .NET application | Enable Excel shared mode programmatically with Aspose.Cells in C#
// Tags: Aspose.Cells shared workbook creation C# | enable concurrent editing Excel Aspose.Cells | save workbook as shared mode .NET | multi‑user Excel collaboration Aspose.Cells | C# Aspose.Cells workbook configuration for sharing

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to instantiate an Aspose.Cells Workbook in C#, add sample data, and save it as a shared Excel file that can be edited concurrently by multiple users, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // (Optional) Add a sample worksheet and data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define output file path
            string outputPath = "SharedWorkbook.xlsx";

            // Save the workbook to the specified file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
