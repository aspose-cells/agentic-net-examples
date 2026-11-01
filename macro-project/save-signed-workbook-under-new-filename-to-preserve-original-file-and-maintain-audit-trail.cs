// Title: Create an audit‑copy of a digitally signed Excel workbook using Aspose.Cells for .NET
// Common Searches: how to copy a signed Excel file to a new name with Aspose.Cells C# | preserve digital signatures when saving an Excel workbook under a different filename in .NET | Aspose.Cells save workbook as audit copy without removing existing signatures | C# duplicate signed workbook while keeping original unchanged using Aspose.Cells | create audit trail for signed Excel workbook by saving with new filename
// Tags: save signed workbook as new file Aspose.Cells | retain existing digital signatures .NET Excel | audit copy generation for signed Excel workbook | Aspose.Cells SaveFormat.Xlsx usage | copy workbook while preserving signatures

using Aspose.Cells;
using System;

// The example loads a digitally signed Excel file (SignedWorkbook.xlsx) with Aspose.Cells, then saves it to a different filename (SignedWorkbook_AuditCopy.xlsx) using SaveFormat.Xlsx, ensuring the original signatures remain intact and providing an audit‑ready copy of the workbook.
class Program
{
    static void Main()
    {
        // Original signed workbook file
        string originalFile = "SignedWorkbook.xlsx";

        // New file name to preserve the original and maintain an audit trail
        string auditFile = "SignedWorkbook_AuditCopy.xlsx";

        // Load the workbook; existing digital signatures are retained automatically
        Workbook workbook = new Workbook(originalFile);

        // Save the workbook under the new filename
        workbook.Save(auditFile, SaveFormat.Xlsx);
    }
}
