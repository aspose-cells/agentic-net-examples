// Title: Add Microsoft Scripting Runtime COM reference to a VBA project with Aspose.Cells (.NET)
// Description: Creates a macro‑enabled workbook, accesses its VbaProject, registers the Microsoft Scripting Runtime (scrrun.dll) via AddRegisteredReference, and saves the file as .xlsm using Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA reference | C# add COM library to VBA project | Microsoft Scripting Runtime | AddRegisteredReference | scrrun.dll | macro-enabled workbook
// Common Searches: Aspose.Cells add Scripting Runtime reference C# | How to register COM library in Excel VBA project programmatically | AddRegisteredReference example for scrrun.dll | Create .xlsm with VBA references using Aspose.Cells | Add COM reference to VBA project without opening Excel
// Developer Intent: Programmatically add the Microsoft Scripting Runtime COM library to the VBA project of a macro‑enabled workbook using Aspose.Cells for .NET.
// Use Cases: Enable Dictionary and FileSystemObject objects in VBA macros by pre‑adding the Scripting Runtime reference. | Distribute template workbooks that rely on external COM components without requiring manual user setup. | Automate generation of macro‑enabled files that need early‑bound access to scripting objects.
// AI Prompts: Write C# code with Aspose.Cells that adds the Microsoft Scripting Runtime reference to a workbook's VBA project and saves it as .xlsm. | Explain the components of the libid string used in VbaProject.References.AddRegisteredReference for a COM library. | Show how to confirm that the Scripting Runtime reference appears in Excel's VBA editor after opening the generated workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    // Creates a macro‑enabled workbook, accesses its VbaProject, registers the Microsoft Scripting Runtime (scrrun.dll) via AddRegisteredReference, and saves the file as .xlsm using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled format will be used on save)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to the Microsoft Scripting Runtime COM library
            // Libid format: *\G{<GUID>}#<Version>#0#<Path>#<Description>
            vbaProject.References.AddRegisteredReference(
                "Scripting", 
                "*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime");

            // Save the workbook as a macro-enabled file
            workbook.Save("WorkbookWithScriptingReference.xlsm", SaveFormat.Xlsm);
        }
    }
}
