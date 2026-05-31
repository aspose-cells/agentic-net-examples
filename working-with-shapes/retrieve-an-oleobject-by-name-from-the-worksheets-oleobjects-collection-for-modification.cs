using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RetrieveAndModifyOleObject
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Name of the OleObject to retrieve
        string oleObjectName = "MyOleObject";

        // Locate the OleObject by its Name property
        OleObject oleToModify = null;
        foreach (OleObject ole in worksheet.OleObjects)
        {
            if (ole.Name == oleObjectName)
            {
                oleToModify = ole;
                break;
            }
        }

        // If the OleObject is found, perform modifications
        if (oleToModify != null)
        {
            // Example modifications
            oleToModify.Label = "UpdatedLabel";
            oleToModify.AutoUpdate = false; // Disable automatic updates
            // Additional property changes can be added here
        }
        else
        {
            Console.WriteLine($"OleObject with name '{oleObjectName}' not found.");
        }

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}