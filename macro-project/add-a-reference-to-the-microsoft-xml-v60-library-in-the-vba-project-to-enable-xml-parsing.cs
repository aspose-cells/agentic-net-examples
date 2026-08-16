// Title: Add MSXML6 Reference to a VBA Project in an Excel Workbook with Aspose.Cells (C#)
// Description: C# code that creates an in‑memory workbook, accesses its VbaProject, registers the Microsoft XML, v6.0 (MSXML6) automation library via a libid string, and saves the file as a macro‑enabled .xlsm so VBA code can parse XML without manual setup.
// Keywords: Aspose.Cells | VbaProject | AddRegisteredReference | MSXML6 | C# Excel macro | XML parsing VBA | macro‑enabled workbook | Automation library reference | libid string | Windows msxml6.dll
// Common Searches: how to add MSXML6 reference to VBA project using Aspose.Cells | C# add Microsoft XML v6.0 to Excel macro | Aspose.Cells VbaProject AddRegisteredReference example | register automation library in .xlsm workbook | save workbook with VBA reference to MSXML6
// Developer Intent: Programmatically register the Microsoft XML, v6.0 library in a workbook’s VBA project and save it as a macro‑enabled file.
// Use Cases: Generate .xlsm templates that already contain the MSXML6 reference for downstream XML‑driven VBA scripts. | Automate creation of reports that rely on MSXML6 for data import, eliminating manual reference setup. | Deploy Excel workbooks across Windows environments where VBA code must parse XML using the built‑in MSXML6 library.
// AI Prompts: Write C# code using Aspose.Cells to add a registered reference to MSXML6 in a workbook’s VBA project and save it as .xlsm. | Explain how to build the libid string for MSXML6 and use VbaProject.References.AddRegisteredReference to register it. | Show how to verify that the MSXML6 reference was successfully added to a saved macro‑enabled workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// C# code that creates an in‑memory workbook, accesses its VbaProject, registers the Microsoft XML, v6.0 (MSXML6) automation library via a libid string, and saves the file as a macro‑enabled .xlsm so VBA code can parse XML without manual setup.
class AddMsXmlReference
{
    static void Main()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to Microsoft XML, v6.0 (MSXML6) as a registered Automation type library
        // The libid format: *\G{<GUID>}#<Version>#0#<Path>#<Description>
        string libid = "*\\G{88D969C5-F192-11D4-A65F-0040963251E5}#6.0#0#C:\\Windows\\System32\\msxml6.dll#Microsoft XML, v6.0";
        vbaProject.References.AddRegisteredReference("MSXML2", libid);

        // Save the workbook as a macro‑enabled file so the VBA project (with the reference) is retained
        workbook.Save("WorkbookWithMsXmlReference.xlsm", SaveFormat.Xlsm);
    }
}
