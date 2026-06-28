using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project (created automatically for macro-enabled workbooks)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a procedural (standard) VBA module
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Procedural, "ExternalComModule");
        VbaModule module = vbaProject.Modules[moduleIndex];

        // Insert VBA code that calls an external COM library (e.g., user32.dll MessageBox)
        module.Codes =
            "Declare PtrSafe Function MessageBoxA Lib \"user32.dll\" (ByVal hwnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long\n" +
            "Sub ShowMessage()\n" +
            "    Call MessageBoxA(0, \"Hello from COM\", \"Aspose.Cells VBA\", 0)\n" +
            "End Sub";

        // Add a reference to an Automation type library (stdole)
        vbaProject.References.AddRegisteredReference(
            "stdole",
            "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

        // Add a control reference (e.g., Microsoft Forms 2.0 Object Library)
        vbaProject.References.AddControlRefrernce(
            "MSForms",
            "*\\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\\Windows\\system32\\FM20.DLL#Microsoft Forms 2.0 Object Library",
            "twiddledLibid",
            "extendedLibid");

        // Save the workbook as a macro‑enabled file
        workbook.Save("ExternalComReferences.xlsm", SaveFormat.Xlsm);
    }
}