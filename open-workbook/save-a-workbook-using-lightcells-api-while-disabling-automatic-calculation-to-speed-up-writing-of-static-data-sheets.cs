// Title: Save an Aspose.Cells workbook without automatic recalculation in C# for static data sheets
// Common Searches: c# aspocells save workbook without triggering calculation | how to prevent formula recalculation when exporting Excel with Aspose.Cells | improve performance writing static data to Excel using Aspose.Cells .NET | disable workbook calculation before saving with Aspose.Cells C# | fast way to save static Excel sheet using Aspose.Cells in .NET
// Tags: disable automatic calculation Aspose.Cells | save workbook without recalculation C# | Aspose.Cells static data performance | Excel file save optimization Aspose.Cells | Aspose.Cells workbook save without calculation

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, inserts static values into the first worksheet, ensures the output directory exists, and saves the file as XLSX while automatic calculation is disabled, resulting in faster write performance for static data.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Insert static data into the first row
            sheet.Cells[0, 0].PutValue("Static value 1");
            sheet.Cells[0, 1].PutValue(12345);

            // Define output file path
            string outputPath = "StaticDataWorkbook.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (no automatic recalculation will be triggered)
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
