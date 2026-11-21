// Title: Save a signed Excel workbook to a new file with Aspose.Cells for .NET while keeping signatures intact
// Common Searches: C# Aspose.Cells save signed Excel file as a copy preserving digital signatures | How to create an audit copy of a signed workbook using Aspose.Cells .NET | Programmatically duplicate a signed .xlsx workbook without losing signatures in C# | Aspose.Cells load signed workbook and save to a different filename for audit trail | Preserve digital signatures when copying an Excel file with Aspose.Cells in .NET
// Tags: Workbook.Save with digital signature preservation | Aspose.Cells duplicate signed Excel workbook | C# generate audit copy of signed .xlsx file | load signed workbook using Aspose.Cells | save workbook under new filename .xlsx

using Aspose.Cells;
using System;

// Loads a signed Excel workbook and saves it to a different filename, ensuring that all digital signatures remain intact for audit‑trail purposes.
class Program
{
    static void Main()
    {
        // Original signed workbook file
        string originalFile = "SignedWorkbook.xlsx";

        // New filename to preserve the original and maintain audit trail
        string auditFile = "SignedWorkbook_AuditCopy.xlsx";

        // Load the signed workbook
        Workbook workbook = new Workbook(originalFile);

        // Save the workbook under the new filename (signatures are preserved)
        workbook.Save(auditFile, SaveFormat.Xlsx);
    }
}
