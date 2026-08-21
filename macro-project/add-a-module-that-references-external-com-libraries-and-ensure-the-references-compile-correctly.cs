// Title: Add a VBA class module with external COM references using Aspose.Cells for .NET (C#)
// Description: C# example that creates a macro‑enabled workbook, accesses its VbaProject, adds a class module, injects VBA code that uses the FileSystemObject COM object, and registers three kinds of external references – a stdole automation type library, the Microsoft Forms 2.0 control library, and an external VBA add‑in – before saving as .xlsm.
// Keywords: Aspose.Cells | C# | VBA class module | COM reference | macro-enabled workbook | xlsm | AddRegisteredReference | AddControlReference | AddProjectReference | stdole | MSForms | FileSystemObject | VbaProject | Aspose.Cells VBA
// Common Searches: how to add VBA class module with COM reference using Aspose.Cells | Aspose.Cells add registered stdole reference C# | add Microsoft Forms 2.0 library to macro workbook Aspose.Cells | link external VBA add‑in (.xlam) with Aspose.Cells | C# create macro‑enabled workbook that calls FileSystemObject
// Developer Intent: Create a macro‑enabled workbook, insert a VBA class module, and register external COM libraries so the VBA code compiles without errors.
// Use Cases: Automate file‑system operations from VBA by exposing the Scripting.FileSystemObject COM object. | Enable userform controls in macros by adding the Microsoft Forms 2.0 Object Library. | Reuse procedures from an existing VBA add‑in across multiple workbooks via a project reference. | Generate macro‑enabled templates programmatically for enterprise reporting solutions.
// AI Prompts: Generate C# code with Aspose.Cells that adds a VBA class module and registers stdole and MSForms COM libraries. | Show how to verify that added COM references compile successfully when saving an .xlsm file with Aspose.Cells. | Provide error‑handling patterns for missing or unregistered COM libraries while adding references via Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// C# example that creates a macro‑enabled workbook, accesses its VbaProject, adds a class module, injects VBA code that uses the FileSystemObject COM object, and registers three kinds of external references – a stdole automation type library, the Microsoft Forms 2.0 control library, and an external VBA add‑in – before saving as .xlsm.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project (automatically created for macro-enabled workbooks)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a class module named "ExternalComModule"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "ExternalComModule");
        VbaModule vbaModule = vbaProject.Modules[moduleIndex];

        // Insert VBA code that uses an external COM library (example using FileSystemObject)
        vbaModule.Codes = "Sub CallExternal()\n    Dim obj As Object\n    Set obj = CreateObject(\"Scripting.FileSystemObject\")\n    MsgBox obj.GetAbsolutePathName(\".\")\nEnd Sub";

        // Add a registered reference to the stdole Automation type library
        vbaProject.References.AddRegisteredReference(
            "stdole",
            "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

        // Add a control reference (e.g., Microsoft Forms 2.0 Object Library)
        vbaProject.References.AddControlRefrernce(
            "MSForms",
            "*\\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\\Windows\\system32\\FM20.DLL#Microsoft Forms 2.0 Object Library",
            "twiddledLibid",
            "extendedLibid");

        // Add a project reference to an external VBA project (example paths)
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",
            "C:\\AddIns\\MyAddIn.xlam",
            "..\\AddIns\\MyAddIn.xlam");

        // Save the workbook as a macro-enabled file
        workbook.Save("WorkbookWithExternalComReferences.xlsm", SaveFormat.Xlsm);
    }
}
