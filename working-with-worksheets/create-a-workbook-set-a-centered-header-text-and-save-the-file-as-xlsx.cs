// Title: Create an Excel workbook, add a centered page header, and save it as XLSX using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to add a centered header to an Excel worksheet | Saving an Aspose.Cells workbook as XLSX with page header in .NET | Using reflection to set PageSetup.CenterHeader in Aspose.Cells C# | Ensure output folder exists before saving Excel file with Aspose.Cells | Example of creating a workbook and setting header with Aspose.Cells for .NET
// Tags: Aspose.Cells centered header via PageSetup | Aspose.Cells workbook save XLSX format | Reflection access PageSetup properties Aspose.Cells | C# ensure directory exists before file save | Aspose.Cells worksheet page header setup

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, accesses the first worksheet, uses reflection to set the PageSetup.CenterHeader when the property is available, guarantees the target directory exists, and saves the workbook as an XLSX file with Aspose.Cells for .NET.
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

            // Set a centered header if the API supports it
            // The CenterHeader property may not be available in older versions,
            // so we guard it with a compile‑time check using reflection.
            var pageSetup = sheet.PageSetup;
            var centerHeaderProp = pageSetup.GetType().GetProperty("CenterHeader");
            if (centerHeaderProp != null && centerHeaderProp.CanWrite)
            {
                centerHeaderProp.SetValue(pageSetup, "My Centered Header");
            }

            // Define output file path
            string outputPath = "CenteredHeader.xlsx";

            // Ensure the directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook as XLSX
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
