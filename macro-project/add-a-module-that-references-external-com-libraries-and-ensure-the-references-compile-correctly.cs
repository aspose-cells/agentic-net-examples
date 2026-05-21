using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsComReferenceDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (macro-enabled format will be used on save)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // ------------------------------------------------------------
            // Add a reference to an external COM (Automation) library.
            // Here we reference the standard OLE Automation library (stdole).
            // ------------------------------------------------------------
            vbaProject.References.AddRegisteredReference(
                "stdole",
                "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

            // ------------------------------------------------------------
            // Optionally, add a control reference (e.g., Microsoft Forms 2.0).
            // This demonstrates adding a twiddled type library reference.
            // ------------------------------------------------------------
            vbaProject.References.AddControlRefrernce(
                "MSForms",
                "*\\G{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0#C:\\Windows\\system32\\FM20.DLL#Microsoft Forms 2.0 Object Library",
                "twiddledLibid_placeholder",
                "extendedLibid_placeholder");

            // ------------------------------------------------------------
            // Add a new VBA module to the project.
            // The module type is Class, and we give it a meaningful name.
            // ------------------------------------------------------------
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "ComHelperModule");
            VbaModule module = vbaProject.Modules[moduleIndex];

            // Insert VBA code that utilizes the referenced COM library.
            // Example: use the stdole library to create a picture object.
            module.Codes = @"
Public Sub ShowStdOleMessage()
    Dim pic As stdole.StdPicture
    MsgBox ""COM library reference is working!""
End Sub
";

            // Save the workbook as a macro-enabled file so that the VBA project is retained.
            workbook.Save("ComReferenceDemo.xlsm", SaveFormat.Xlsm);
        }
    }
}