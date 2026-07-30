// Title: Password‑protect an Excel worksheet and export it to CSV with Aspose.Cells for .NET
// Description: Shows how to create a workbook, apply full worksheet protection using a password, save the sheet as a CSV file (the protection does not alter the exported data), then remove protection and save the workbook as XLSX.
// Keywords: Aspose.Cells | C# | worksheet protection | password protection | CSV export | SaveFormat.Csv | SaveFormat.Xlsx | Protect | Unprotect | ProtectionType.All
// Common Searches: Aspose.Cells protect worksheet and export to CSV | Does worksheet protection affect CSV output in Aspose.Cells | C# export protected sheet to CSV using Aspose.Cells | How to unprotect a worksheet after CSV export with Aspose.Cells | Save protected Excel sheet as CSV .NET
// Developer Intent: Apply a password to secure a worksheet, then generate a CSV file that contains the original data unchanged, and finally remove the protection for further Excel processing.
// Use Cases: Secure a workbook before distribution while still providing a CSV report for downstream analysis. | Automate a pipeline that locks worksheets, creates CSV extracts for data pipelines, and later releases the lock for archival storage. | Generate CSV snapshots of protected financial models without exposing the underlying Excel file's edit capabilities.
// AI Prompts: Provide C# code using Aspose.Cells to protect a worksheet with ProtectionType.All, export it to CSV without the protection affecting the output, then unprotect and save as XLSX. | Explain why worksheet protection does not change CSV export results in Aspose.Cells for .NET. | Create a step‑by‑step guide to lock a sheet, export to CSV, and later unlock and save the workbook in Excel format using Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create a workbook, apply full worksheet protection using a password, save the sheet as a CSV file (the protection does not alter the exported data), then remove protection and save the workbook as XLSX.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);

        // Protect the worksheet with a password (all protection types)
        sheet.Protect(ProtectionType.All, "pwd123", null);

        // Export the worksheet data to CSV.
        // Worksheet protection does not affect the exported values.
        workbook.Save("ProtectedWorksheet.csv", SaveFormat.Csv);

        // (Optional) Unprotect and save as Excel to verify protection removal
        sheet.Unprotect("pwd123");
        workbook.Save("UnprotectedWorksheet.xlsx", SaveFormat.Xlsx);
    }
}
