// Title: Create an audit‑trail copy of a digitally signed Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load a signed .xlsx workbook with Aspose.Cells and save it under a new filename to keep the original unchanged. | Generate an audit copy of a digitally signed Excel file in C# while preserving its digital signatures. | Duplicate a signed workbook to a separate file using Aspose.Cells without affecting the original signature.
// Common Searches: asp.net how to duplicate a signed Excel file with Aspose.Cells preserving the signature | c# save signed workbook as new file using Aspose.Cells for audit purposes | preserve digital signature when copying an .xlsx workbook with Aspose.Cells | create audit trail copy of a digitally signed Excel workbook in .NET
// Tags: Aspose.Cells save workbook with new filename | duplicate signed Excel workbook C# | preserve digital signature on Excel save | audit trail copy of signed .xlsx | load signed workbook Aspose.Cells

using Aspose.Cells;
using System;

// Loads a digitally signed Excel workbook (SignedWorkbook.xlsx) with Aspose.Cells and saves it as a new file (SignedWorkbook_AuditCopy.xlsx) using SaveFormat.Xlsx, preserving the original file and its digital signatures for audit tracking.
class Program
{
    static void Main()
    {
        // Original signed workbook file
        string originalFile = "SignedWorkbook.xlsx";

        // New file name to preserve the original and create an audit trail
        string auditFile = "SignedWorkbook_AuditCopy.xlsx";

        // Load the workbook; digital signatures are retained automatically
        Workbook workbook = new Workbook(originalFile);

        // Save the workbook under the new name, keeping the original unchanged
        workbook.Save(auditFile, SaveFormat.Xlsx);
    }
}
