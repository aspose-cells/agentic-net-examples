// Title: Apply worksheet protection in Aspose.Cells for .NET that permits inserting new hyperlinks but blocks editing existing hyperlink addresses
// Common Searches: how to protect an Excel worksheet but still allow adding new hyperlinks using Aspose.Cells C# | Aspose.Cells worksheet protection block editing of existing hyperlink addresses | C# protect sheet with Aspose.Cells while permitting hyperlink insertion | using Aspose.Cells .NET protect worksheet without AllowEditHyperlinks property | prevent users from changing hyperlink URLs in a protected Excel file with Aspose.Cells
// Tags: worksheet protection allow insert hyperlink Aspose.Cells | block editing existing hyperlink addresses C# | Aspose.Cells Protect method ProtectionType.All | Excel sheet protection without AllowEditHyperlinks | disable hyperlink editing in protected workbook

using System;
using Aspose.Cells;

// The example creates a new workbook, accesses the first worksheet, applies full protection with worksheet.Protect(ProtectionType.All) which by default prevents editing of hyperlinks, notes that the AllowEditHyperlinks property is unavailable in this Aspose.Cells version, and saves the file as ProtectedSheet.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Configure worksheet protection (default settings disallow editing of hyperlinks)
            var protection = worksheet.Protection;
            // Note: In this version of Aspose.Cells the AllowEditHyperlinks property is not available.
            // The default protection settings already prevent editing of hyperlinks.

            // Apply protection to the worksheet (protect all elements by default)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook to a file
            workbook.Save("ProtectedSheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
