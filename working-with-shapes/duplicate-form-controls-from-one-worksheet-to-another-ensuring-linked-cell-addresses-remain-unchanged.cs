// Title: Copy a worksheet with form controls while keeping linked cells unchanged – Aspose.Cells C#
// Description: Shows how to load a workbook, clone a sheet that contains check boxes, drop‑downs or other form controls using Workbook.Worksheets.AddCopy, and save the new file. The cloned sheet retains the original controls’ cell bindings because no copy options modify them.
// Keywords: Aspose.Cells AddCopy | duplicate worksheet C# | form controls copy | linked cell reference | preserve control bindings | copy sheet drawing objects | Aspose.Cells .NET example | clone worksheet with controls
// Common Searches: Aspose.Cells copy worksheet with form controls | how to keep linked cells when duplicating a sheet in .NET | clone Excel sheet containing checkboxes using Aspose.Cells | preserve form control references during worksheet copy | AddCopy vs CopyOptions for form controls
// Developer Intent: Create an exact replica of a worksheet that includes form controls, ensuring the controls continue to point to their original cells.
// Use Cases: Generate per‑user copies of a template sheet that contains check boxes and drop‑downs without breaking the cell links. | Automate the creation of monthly report tabs that share the same control layout and bindings. | Batch‑process multiple sheets with embedded radio buttons while maintaining their data connections.
// AI Prompts: Write C# code with Aspose.Cells to duplicate a worksheet that has form controls and retain the original linked cell addresses. | Provide an example that loops through several worksheets, copies each one with its controls, and preserves all cell references. | Explain the impact of using AddCopy versus AddCopy with CopyOptions on form control links in Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to load a workbook, clone a sheet that contains check boxes, drop‑downs or other form controls using Workbook.Worksheets.AddCopy, and save the new file. The cloned sheet retains the original controls’ cell bindings because no copy options modify them.
class DuplicateFormControls
{
    static void Main()
    {
        // Load the workbook that contains the worksheet with form controls
        Workbook workbook = new Workbook("source.xlsx");

        // Assume the first worksheet holds the form controls to duplicate
        Worksheet sourceSheet = workbook.Worksheets[0];
        string sourceSheetName = sourceSheet.Name;

        // Duplicate the worksheet within the same workbook.
        // AddCopy copies the worksheet contents, formats, and drawing objects (including form controls).
        // The linked cell addresses of the controls remain unchanged because we do not alter copy options.
        int copiedIndex = workbook.Worksheets.AddCopy(sourceSheetName);
        Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

        // Optionally give the copied sheet a distinct name
        copiedSheet.Name = sourceSheetName + "_Copy";

        // Save the workbook with the duplicated worksheet
        workbook.Save("output.xlsx");
    }
}
