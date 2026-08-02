// Title: Protect an Excel worksheet with Aspose.Cells for .NET and export to CSV
// Description: Shows how to apply full worksheet protection (without a password) using Aspose.Cells for .NET and then save the workbook as a CSV file, confirming that protection does not change the exported data.
// Keywords: Aspose.Cells worksheet protection | export protected sheet to CSV | Aspose.Cells .NET CSV export | Excel sheet protection Aspose | save workbook as CSV Aspose.Cells | ProtectionType.All Aspose.Cells | worksheet protection impact on CSV
// Common Searches: Aspose.Cells protect worksheet and export CSV | Does worksheet protection affect CSV output in Aspose.Cells | Save protected Excel sheet as CSV using C# | Export data from a protected worksheet with Aspose.Cells | How to keep cell values when saving protected sheet to CSV
// Developer Intent: Apply worksheet protection and generate a CSV file that contains the original cell values unchanged.
// Use Cases: Secure an Excel file for internal distribution while still producing CSV reports for downstream systems. | Create a read‑only workbook and automate CSV extraction for data migration without disabling protection. | Programmatically lock a sheet before archiving, then export its contents for analytics or backup.
// AI Prompts: Write C# code with Aspose.Cells that protects a worksheet with a password and exports it to CSV, ensuring the password does not block the export. | Demonstrate how to protect only formatting or editing features of a sheet and then save it as CSV using Aspose.Cells. | Explain the relationship between worksheet protection settings and different save formats in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to apply full worksheet protection (without a password) using Aspose.Cells for .NET and then save the workbook as a CSV file, confirming that protection does not change the exported data.
class ProtectAndExportCsv
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Fill some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);

        // Protect the worksheet (all protection types, no password)
        sheet.Protect(ProtectionType.All);

        // Export the worksheet data to CSV; protection does not affect the exported values
        workbook.Save("ProtectedData.csv", SaveFormat.Csv);

        // Optional: display protection status
        Console.WriteLine("Worksheet protected: " + sheet.IsProtected);
    }
}
