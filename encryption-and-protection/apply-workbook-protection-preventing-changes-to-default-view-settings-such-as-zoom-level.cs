// Title: C# Aspose.Cells – Password‑protect Workbook Window (zoom) settings
// Description: Creates a new Workbook, locks the default view (window) using ProtectionType.Windows with a password, and saves it as ProtectedDefaultView.xlsx.
// Keywords: Aspose.Cells C# protect window | Excel zoom lock programmatically | Workbook view protection .NET | ProtectionType.Windows password | Secure default view settings
// Common Searches: Aspose.Cells lock workbook zoom level | C# protect Excel window settings with password | Prevent view changes in Excel using Aspose | How to disable zoom editing in a .NET workbook | Password protect Excel default view options
// Developer Intent: Apply a password to the workbook’s window settings so the default view (e.g., zoom) cannot be altered by end users.
// Use Cases: Distribute a template that must always open at a predefined zoom for brand consistency. | Deliver a financial dashboard where layout integrity is critical and view changes are prohibited. | Enforce corporate policy that only authorized staff can modify workbook display options.
// AI Prompts: Write C# code with Aspose.Cells to lock the workbook window (zoom) using a custom password and save the file. | Explain how to programmatically remove the window protection from a workbook in Aspose.Cells. | Show how to protect both the workbook structure and its window settings in one Aspose.Cells call.

using Aspose.Cells;

// Creates a new Workbook, locks the default view (window) using ProtectionType.Windows with a password, and saves it as ProtectedDefaultView.xlsx.
class ProtectWorkbookDefaultView
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook's window (default view settings such as zoom level) with a password
        workbook.Protect(ProtectionType.Windows, "pwd123");

        // Save the protected workbook
        workbook.Save("ProtectedDefaultView.xlsx");
    }
}
