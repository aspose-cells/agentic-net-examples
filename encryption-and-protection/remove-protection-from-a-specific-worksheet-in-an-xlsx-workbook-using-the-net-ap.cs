using System;
using Aspose.Cells;

class RemoveWorksheetProtection
{
    static void Main()
    {
        // Load the existing workbook that contains a protected worksheet
        Workbook workbook = new Workbook("protected_input.xlsx");

        // Access the specific worksheet to be unprotected (by name or index)
        Worksheet worksheet = workbook.Worksheets["Sheet1"]; // replace with actual sheet name or use workbook.Worksheets[0]

        // Unprotect the worksheet.
        // If the worksheet was protected with a password, provide it; otherwise call Unprotect() without arguments.
        worksheet.Unprotect("yourPassword"); // replace with the actual password or use worksheet.Unprotect();

        // Save the workbook after removing protection
        workbook.Save("unprotected_output.xlsx");
    }
}