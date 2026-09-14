// Title: Use Aspose.Cells for .NET to password‑protect workbook structure and verify that sheet removal is prevented
// AI Prompts: Write C# code that creates a workbook, adds worksheets, applies structure protection with a password using Aspose.Cells, then attempts to delete a sheet and captures the resulting exception. | Show how to handle the error raised when trying to remove a worksheet from a structure‑protected workbook and then save the file as an .xlsx.
// Common Searches: aspnet protect Excel workbook structure with password using Aspose.Cells | c# prevent worksheet deletion after applying workbook protection Aspose.Cells | how to handle error when removing sheet from a protected workbook in Aspose.Cells | example of using workbook.Protect(ProtectionType.Structure) in C# | verify that structure protection blocks sheet removal with Aspose.Cells
// Tags: structure protection Aspose.Cells C# | password protect workbook Aspose.Cells | block worksheet deletion Aspose.Cells | handle protected sheet removal error C# | save protected workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The sample creates a workbook with two sheets, applies password‑based structure protection via Aspose.Cells, attempts to delete the first sheet (which triggers an exception), catches and logs the expected error, and finally saves the protected workbook as ProtectedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Sheet1";
            workbook.Worksheets.Add("Sheet2");

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "pwd123");

            // Attempt to delete a worksheet; this should be blocked by the protection
            try
            {
                workbook.Worksheets.RemoveAt(0);
                Console.WriteLine("Worksheet deleted (unexpected).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Deletion blocked as expected: " + ex.Message);
            }

            // Save the workbook
            workbook.Save("ProtectedWorkbook.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
