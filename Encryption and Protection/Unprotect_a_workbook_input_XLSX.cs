using System;
using Aspose.Cells;
using System.IO;

class Program
{
    static void Main()
    {
        const string password = "yourPassword";
        string protectedPath = Path.Combine(Directory.GetCurrentDirectory(), "protected_input.xlsx");
        string unprotectedPath = Path.Combine(Directory.GetCurrentDirectory(), "unprotected_output.xlsx");

        // Create a sample workbook and protect it if it doesn't already exist
        if (!File.Exists(protectedPath))
        {
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
            wb.Protect(ProtectionType.Structure, password);
            wb.Save(protectedPath);
        }

        // Load the protected workbook
        Workbook workbook = new Workbook(protectedPath);

        // Unprotect the workbook
        workbook.Unprotect(password);

        // Save the unprotected workbook
        workbook.Save(unprotectedPath);
    }
}