// Title: Copy a VBA UserForm with Chart Control Between XLSM Workbooks Using Aspose.Cells for .NET
// Description: Demonstrates how to load a macro‑enabled source workbook, create an empty destination workbook, and transfer the entire VBA project—including UserForms that contain chart controls—using the Aspose.Cells VbaProject.Copy method, then save the result as a new XLSM file.
// Keywords: Aspose.Cells | C# | VbaProject.Copy | copy VBA UserForm | chart control in UserForm | macro-enabled workbook | XLSM transfer | Excel VBA project copy | automate Excel forms | Aspose.Cells example
// Common Searches: copy UserForm with chart from one XLSM to another Aspose.Cells | Aspose.Cells VbaProject.Copy example C# | transfer VBA UserForm between macro‑enabled workbooks | how to move Excel UserForm chart control using Aspose.Cells | duplicate VBA project including forms with Aspose.Cells
// Developer Intent: Transfer a VBA UserForm that contains a chart control from a source XLSM workbook to a destination XLSM workbook using Aspose.Cells for .NET.
// Use Cases: Clone a template workbook that includes custom UI forms with embedded charts for batch report generation. | Migrate legacy macro‑enabled spreadsheets to new files while preserving all VBA forms and their visual components. | Create a clean copy of a workbook for external distribution that still contains the original VBA forms for internal editing.
// AI Prompts: Generate C# code to copy only a selected UserForm (excluding other VBA modules) between XLSM files with Aspose.Cells. | Explain how to handle missing VbaProject objects when using VbaProject.Copy and suggest fallback strategies. | Show how to programmatically verify that a chart control inside a transferred UserForm remains functional after the copy operation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Demonstrates how to load a macro‑enabled source workbook, create an empty destination workbook, and transfer the entire VBA project—including UserForms that contain chart controls—using the Aspose.Cells VbaProject.Copy method, then save the result as a new XLSM file.
class TransferUserFormWithChart
{
    static void Main()
    {
        string sourcePath = "SourceWithUserForm.xlsm";
        string destPath = "DestinationWithUserForm.xlsm";

        // Verify source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the source workbook (must be macro-enabled)
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Copy the VBA project (including UserForms and controls) from source to destination
            if (sourceWorkbook.VbaProject != null && destinationWorkbook.VbaProject != null)
            {
                destinationWorkbook.VbaProject.Copy(sourceWorkbook.VbaProject);
            }
            else
            {
                Console.WriteLine("VBA project is missing in source or destination workbook.");
            }

            // Save the destination workbook as a macro-enabled file
            destinationWorkbook.Save(destPath, SaveFormat.Xlsm);
            Console.WriteLine($"Destination workbook saved successfully: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
