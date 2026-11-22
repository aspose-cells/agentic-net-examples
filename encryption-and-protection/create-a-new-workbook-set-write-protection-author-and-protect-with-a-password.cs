// Title: Create a new Excel workbook and apply write‑protection with author and password using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set workbook write protection author and password | How to protect a newly created Excel file with author name using Aspose.Cells .NET | Set write protection author on Excel workbook programmatically with Aspose.Cells | Save an Excel workbook with password protection via Aspose.Cells in C# | Apply write protection settings to a workbook before saving with Aspose.Cells
// Tags: Aspose.Cells workbook write protection API | Excel file password protection using Aspose.Cells | set author for workbook protection .NET | protect new workbook with write protection settings | Aspose.Cells Settings.WriteProtection usage

using System;
using Aspose.Cells;

// The example creates a new Workbook, configures the Settings.WriteProtection object with an author name and password, and saves the file as a password‑protected Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Configure write protection (author and password)
            WriteProtection protection = workbook.Settings.WriteProtection;
            protection.Author = "John Doe";
            protection.Password = "MySecurePassword"; // Set password directly

            // Save the protected workbook
            workbook.Save("ProtectedWorkbook.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
