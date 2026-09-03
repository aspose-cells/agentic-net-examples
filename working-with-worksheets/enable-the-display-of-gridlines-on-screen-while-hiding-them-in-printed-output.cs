// Title: Enable on-screen gridlines but suppress them in printed output with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that sets Worksheet.IsGridlinesVisible to true and PageSetup.PrintGridlines to false, then saves the workbook as XLSX. | Update an existing Aspose.Cells workbook so that gridlines appear in the Excel UI while they are omitted from the printed page, preserving all other worksheet data.
// Common Searches: aspocells C# show gridlines in Excel view but not when printing | how to hide gridlines in printed Excel file using Aspose.Cells | set Worksheet.IsGridlinesVisible and PageSetup.PrintGridlines properties in .NET | save workbook with screen gridlines only Aspose.Cells example
// Tags: gridlines visibility on screen Aspose.Cells | suppress printed gridlines Aspose.Cells | Worksheet gridlines visibility C# example | disable print gridlines Aspose.Cells | save workbook with custom gridline settings Aspose.Cells | Excel UI gridlines vs print gridlines Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, makes gridlines visible in the Excel UI via Worksheet.IsGridlinesVisible, disables gridlines for printed output using PageSetup.PrintGridlines, and saves the file as GridlinesDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (using the create rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Show gridlines on screen
            sheet.IsGridlinesVisible = true;

            // Hide gridlines when printing
            sheet.PageSetup.PrintGridlines = false;

            // Define output file path
            string outputPath = "GridlinesDemo.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (using the save rule)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
