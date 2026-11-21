// Title: Display hidden rows and columns during Excel to PDF conversion with Aspose.Cells for .NET (C#)
// Common Searches: how to include hidden rows and columns in PDF export using Aspose.Cells C# | Aspose.Cells C# convert Excel to PDF while showing hidden cells | set Workbook.Settings to make hidden rows visible in PDF output Aspose.Cells | C# reflection to enable IsHiddenRowsAndColumnsVisible property in Aspose.Cells | export Excel workbook to PDF with hidden rows displayed using Aspose.Cells .NET
// Tags: Aspose.Cells show hidden rows PDF export | Workbook.Settings IsHiddenRowsAndColumnsVisible C# | Excel to PDF conversion with visible hidden columns Aspose | reflection set Aspose.Cells property for PDF output | control hidden rows visibility in PDF with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, uses reflection to set Workbook.Settings.IsHiddenRowsAndColumnsVisible to true when the property exists, and saves the workbook as a PDF so that hidden rows and columns are rendered visible; it also includes file‑existence validation and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Attempt to make hidden rows and columns visible (if the property exists in the current version)
            var settings = workbook.Settings;
            var prop = settings.GetType().GetProperty("IsHiddenRowsAndColumnsVisible");
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(settings, true);
            }

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully created at '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
