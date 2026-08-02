// Title: Lock Excel Workbook View Settings (Zoom, Window) with Aspose.Cells for .NET
// Description: Shows how to use Aspose.Cells' Protect method with ProtectionType.Windows and a password to block changes to a workbook's window options—such as zoom level and view layout—and then save the protected file.
// Keywords: Aspose.Cells Protect Windows | C# lock Excel zoom level | prevent view changes Aspose.Cells | Workbook.Protect password | Excel workbook window protection | Aspose.Cells default view lock | C# Excel protection types | ProtectionType.Windows example
// Common Searches: Aspose.Cells protect zoom level | C# lock Excel view settings with password | prevent users from changing Excel window options using Aspose.Cells | Workbook.Protect Windows example C# | secure Excel default view Aspose.Cells
// Developer Intent: The developer wants to apply workbook protection that blocks modifications to default view settings such as zoom level and window layout.
// Use Cases: Distribute a template where the visual layout must remain identical for every recipient. | Publish a financial report that always opens at a predefined zoom level to preserve formatting. | Enforce corporate presentation standards by locking window options in compliance‑focused workbooks.
// AI Prompts: Provide C# code that protects only the window settings of an Aspose.Cells workbook with a password and later unprotects it. | Explain all ProtectionType values in Aspose.Cells and illustrate which one locks default view settings. | Show a combined example that applies both Windows and Structure protection to a workbook using Aspose.Cells.

using Aspose.Cells;

// Shows how to use Aspose.Cells' Protect method with ProtectionType.Windows and a password to block changes to a workbook's window options—such as zoom level and view layout—and then save the protected file.
class ProtectWorkbookDefaultView
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Protect the workbook's window settings (e.g., zoom level, view options)
        workbook.Protect(ProtectionType.Windows, "pwd123");

        // Save the protected workbook
        workbook.Save("ProtectedDefaultView.xlsx");
    }
}
