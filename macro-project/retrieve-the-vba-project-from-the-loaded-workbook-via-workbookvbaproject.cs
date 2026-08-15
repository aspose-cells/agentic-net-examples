// Title: C# – Retrieve VBA Project from a Macro‑Enabled Workbook with Aspose.Cells
// Description: This example demonstrates how to load an .xlsm file using Aspose.Cells for .NET, access its VBA project via the Workbook.VbaProject property, verify its existence, and output key properties such as Name, IsSigned and IsProtected before saving the workbook unchanged.
// Keywords: Aspose.Cells VBA project | Workbook.VbaProject C# | read VBA metadata Aspose | macro‑enabled workbook .xlsm | check VBA project signed | detect protected VBA project | Aspose.Cells .NET example
// Common Searches: how to get VBA project from xlsm using Aspose.Cells | Aspose.Cells retrieve VBA project name | C# check if workbook contains VBA project | read VBA project properties with Aspose.Cells | access Workbook.VbaProject property
// Developer Intent: Extract the VBA project from a loaded workbook and read its basic attributes.
// Use Cases: Confirm that an incoming .xlsm file includes a VBA project before further processing. | Log VBA project details (name, signing status, protection flag) for compliance auditing. | Validate that a workbook’s VBA project is signed and not password‑protected prior to distribution.
// AI Prompts: Write C# code with Aspose.Cells that extracts the VBA project from a workbook and enumerates all its modules. | Show how to rename a VBA project and save the workbook as a macro‑enabled file using Aspose.Cells. | Explain how to detect whether a VBA project is password‑protected with the Aspose.Cells VbaProject API.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

// This example demonstrates how to load an .xlsm file using Aspose.Cells for .NET, access its VBA project via the Workbook.VbaProject property, verify its existence, and output key properties such as Name, IsSigned and IsProtected before saving the workbook unchanged.
class RetrieveVbaProject
{
    static void Main()
    {
        // Load an existing macro-enabled workbook
        Workbook workbook = new Workbook("input.xlsm");

        // Retrieve the VBA project from the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project exists and display some of its properties
        if (vbaProject != null)
        {
            Console.WriteLine("VBA Project Name: " + vbaProject.Name);
            Console.WriteLine("Is Signed: " + vbaProject.IsSigned);
            Console.WriteLine("Is Protected: " + vbaProject.IsProtected);
        }
        else
        {
            Console.WriteLine("No VBA project found in the workbook.");
        }

        // Save the workbook (unchanged) as a macro-enabled file
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
