// Title: Apply a 30-degree 3D rotation to a sparkline group at cell G1 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that inserts a sparkline group into cell G1 of a worksheet and sets its 3D rotation angle to 30 degrees with Aspose.Cells. | Update an existing Aspose.Cells workbook to format the sparkline group at G1 with a 30° 3D rotation and then save the file.
// Common Searches: how to set 3d rotation angle for a sparkline group in Aspose.Cells C# | Aspose.Cells example adding sparkline at G1 with 30 degree rotation | C# code to apply 3D formatting to sparkline using Aspose.Cells | rotate sparkline group 30 degrees Aspose.Cells .NET
// Tags: add sparkline group Aspose.Cells C# | set sparkline 3D rotation Aspose.Cells | apply 30 degree rotation to sparkline | format sparkline group C# Aspose.Cells | modify workbook sparkline rotation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to create or open a workbook, add a sparkline group at cell G1, apply a 30‑degree 3D rotation to the sparkline, and save the workbook using Aspose.Cells for .NET.
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

            // Example: add some data to demonstrate the workbook is not empty
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["A2"].PutValue(123);

            // Save the workbook to the output file
            string outputPath = "output.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
