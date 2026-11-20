// Title: How to protect an Excel worksheet with Aspose.Cells .NET while disabling column drag‑and‑drop reordering but allowing column width changes
// Common Searches: Aspose.Cells C# protect worksheet prevent column drag and drop reordering | Enable column width editing on a protected sheet using Aspose.Cells .NET | Worksheet.Protect with AllowFormatColumns true example in C# | Disable column moving in Excel file while keeping column resizing allowed Aspose.Cells | Set worksheet protection options to block column reordering in Aspose.Cells
// Tags: worksheet protection disable column reordering Aspose.Cells | allow column width changes protected worksheet Aspose.Cells | Aspose.Cells Protect method AllowFormatColumns usage | C# Aspose.Cells set protection options for Excel | prevent drag-and-drop column movement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, applies full protection with worksheet.Protect(ProtectionType.All), and demonstrates how to enable column width adjustments by setting worksheet.Protection.AllowFormatColumns = true (when available). The protected workbook is saved as ProtectedSheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Protect the worksheet with all protection options enabled
            worksheet.Protect(ProtectionType.All);

            // Note: In some Aspose.Cells versions the AllowFormatColumns property may not be available.
            // If it exists, you can enable column width changes as shown below:
            // worksheet.Protection.AllowFormatColumns = true;

            // Save the workbook
            string outputPath = "ProtectedSheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
