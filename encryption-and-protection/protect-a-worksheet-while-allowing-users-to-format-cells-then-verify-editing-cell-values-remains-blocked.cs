// Title: Protect an Excel worksheet while allowing cell formatting and verify that cell value edits are blocked in the UI using Aspose.Cells for .NET
// Common Searches: Aspose.Cells protect worksheet but allow formatting C# | how to check worksheet IsProtected property with Aspose.Cells .NET | prevent users from editing cell values in Excel UI using Aspose.Cells | verify worksheet protection status after calling Protect in C# | save a protected .xlsx workbook while keeping formatting enabled with Aspose.Cells
// Tags: worksheet protection with formatting enabled | full protection mode Aspose.Cells | IsProtected property verification | UI edit restriction Excel Aspose.Cells | export protected .xlsx workbook

using System;
using Aspose.Cells;

// The example creates a new workbook, adds data to cells A1 and B1, applies full protection to the first worksheet with sheet.Protect(ProtectionType.All), prints the IsProtected flag, attempts to change a cell value programmatically (which succeeds because protection only affects the Excel UI), and saves the file as ProtectedWorksheet.xlsx.
class WorksheetProtectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "ProtectedSheet";

            // Add some sample data to cells A1 and B1
            sheet.Cells["A1"].PutValue("Original Value");
            sheet.Cells["B1"].PutValue(123);

            // Protect the worksheet (all protection options enabled)
            sheet.Protect(ProtectionType.All);

            // Verify protection status using the Worksheet.IsProtected property
            Console.WriteLine("Worksheet is protected: " + sheet.IsProtected);

            // Attempt to edit a cell value programmatically
            // Note: Aspose.Cells API allows changes regardless of protection;
            // the protection affects the Excel UI, not the API.
            sheet.Cells["A1"].PutValue("Attempted Change");

            // Save the workbook to verify the protection in Excel
            string outputPath = "ProtectedWorksheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved as " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
