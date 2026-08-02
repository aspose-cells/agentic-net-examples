// Title: C# – Password‑protect a single worksheet with Aspose.Cells and save the workbook
// Description: Loads an existing Excel file using Aspose.Cells, selects a worksheet, sets a password, applies full protection, and saves the workbook with the protected sheet.
// Keywords: Aspose.Cells | C# worksheet protection | Excel password protect sheet | protect specific worksheet Aspose | ProtectionType.All | Aspose.Cells .NET encryption | secure Excel sheet programmatically
// Common Searches: Aspose.Cells protect one worksheet with password C# | How to lock a single Excel sheet using Aspose.Cells .NET | C# code to add password protection to a worksheet in Aspose.Cells | Save workbook after applying worksheet protection Aspose.Cells | Encrypt only one sheet in an existing Excel file with Aspose
// Developer Intent: Add password protection to a chosen worksheet in an Excel workbook and persist the changes using Aspose.Cells for .NET.
// Use Cases: Distribute a financial report while preventing edits to the calculations sheet. | Share a template where the layout must stay unchanged but other sheets remain editable. | Protect confidential data in a single worksheet before sending the file to external partners.
// AI Prompts: Generate C# code that protects multiple worksheets, each with a distinct password, using Aspose.Cells. | Show how to programmatically remove worksheet protection and retrieve the original password with Aspose.Cells for .NET. | Explain how to combine worksheet password protection with cell‑level read‑only settings in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetProtection
{
    // Loads an existing Excel file using Aspose.Cells, selects a worksheet, sets a password, applies full protection, and saves the workbook with the protected sheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the worksheet you want to protect (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name

            // Set the password for the worksheet protection
            worksheet.Protection.Password = "mySecretPassword";

            // Apply protection to the worksheet (protect all aspects)
            worksheet.Protect(ProtectionType.All, "mySecretPassword", null);

            // Save the workbook with the protected worksheet
            workbook.Save("output.xlsx");
        }
    }
}
