// Title: C# – Protect an Aspose.Cells Worksheet with Password “Secure123” and Allow Only Locked‑Cell Selection
// Description: Demonstrates how to create a workbook, set a password, enable selection of locked cells while disabling selection of unlocked cells, apply full protection (ProtectionType.All), and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells worksheet protection C# | protect worksheet password Aspose | allow selecting locked cells only | ProtectionType.All example | C# Excel file security Aspose.Cells | set worksheet protection options programmatically
// Common Searches: Aspose.Cells protect worksheet with password C# | allow only locked cells to be selected Aspose.Cells | disable selection of unlocked cells in Excel using Aspose | C# code for worksheet protection Aspose.Cells | how to set ProtectionType.All in Aspose.Cells
// Developer Intent: Apply password‑protected protection to a worksheet while permitting users to select only the locked cells.
// Use Cases: Distribute a read‑only financial report where users can highlight locked data but cannot edit any cells. | Provide a template that blocks all edits except for a few unlocked input fields, while still allowing navigation to locked sections. | Secure confidential spreadsheets in a corporate portal, ensuring users can view but not modify protected content.
// AI Prompts: Generate C# code with Aspose.Cells to protect a worksheet using password 'Secure123' and enable selection of locked cells only. | Explain how to modify the protection settings to also allow selection of unlocked cells while keeping the worksheet password. | Show a step‑by‑step example for protecting multiple worksheets in a workbook, each with a different password, using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, set a password, enable selection of locked cells while disabling selection of unlocked cells, apply full protection (ProtectionType.All), and save the file using Aspose.Cells for .NET.
public class ProtectWorksheetDemo
{
    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = worksheet.Protection;

        // Allow users to select locked cells only
        protection.AllowSelectingLockedCell = true;
        protection.AllowSelectingUnlockedCell = false; // optional, default is false

        // Set the protection password
        protection.Password = "Secure123";

        // Protect the worksheet with all protection options
        worksheet.Protect(ProtectionType.All);

        // Save the workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            ProtectWorksheetDemo.Run();
            Console.WriteLine("Workbook protected and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
