using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RefreshOleObjects
{
    static void Main()
    {
        // Load the workbook that contains linked OLE objects
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each OLE object in the current worksheet
            foreach (OleObject ole in sheet.OleObjects)
            {
                // Process only OLE objects that are linked to external files
                if (ole.IsLink)
                {
                    // Enable automatic update so the linked object reflects the latest source data
                    ole.AutoUpdate = true;
                }
            }
        }

        // Save the workbook after refreshing the linked OLE objects
        workbook.Save("output.xlsx");
    }
}