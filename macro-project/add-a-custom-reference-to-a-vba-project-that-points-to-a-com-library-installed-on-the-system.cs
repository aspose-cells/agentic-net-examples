// Title: Add a Custom COM Library Reference to a VBA Project in a Macro‑Enabled Workbook (Aspose.Cells for .NET)
// Description: Demonstrates how to create a macro‑enabled workbook, access its VbaProject, and register a COM type library (e.g., stdole) using Aspose.Cells. The example builds the libid string, adds the reference, and saves the file as an .xlsm document.
// Keywords: Aspose.Cells VBA COM reference | add COM library to VbaProject .NET | macro enabled workbook reference stdole | VbaProject References.AddRegisteredReference | early binding COM type library Excel | C# Aspose.Cells VBA automation | global office integration
// Common Searches: how to register a COM library in VBA project with Aspose.Cells | Aspose.Cells add stdole reference C# | add custom COM reference to Excel macro workbook .NET | VbaProject AddRegisteredReference example | save workbook with VBA references using Aspose.Cells
// Developer Intent: Register a system‑installed COM type library in the VBA project of a macro‑enabled Excel file.
// Use Cases: Enable early‑bound OLE Automation (stdole) for generated VBA macros. | Integrate third‑party COM components so VBA code created by Aspose.Cells can instantiate their objects. | Distribute workbooks that automatically link to required type libraries on Windows machines.
// AI Prompts: Write C# code with Aspose.Cells that adds a registered COM reference to a VBA project and includes error handling for missing type libraries. | Show how to enumerate existing VBA references in a workbook, then add a new COM reference using a custom libid string. | Provide a step‑by‑step tutorial for adding a COM library to a macro‑enabled workbook and saving it as .xlsm with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to create a macro‑enabled workbook, access its VbaProject, and register a COM type library (e.g., stdole) using Aspose.Cells. The example builds the libid string, adds the reference, and saves the file as an .xlsm document.
class Program
{
    static void Main()
    {
        // Create a new workbook (macro-enabled)
        Workbook workbook = new Workbook();

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to a COM library (example: stdole)
        // The libid string follows the format:
        // "*\\G{<GUID>}#<Version>#0#<Path>#<Description>"
        string referenceName = "stdole";
        string libid = "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation";

        vbaProject.References.AddRegisteredReference(referenceName, libid);

        // Save the workbook as a macro-enabled file
        workbook.Save("CustomComReference.xlsm", SaveFormat.Xlsm);
    }
}
