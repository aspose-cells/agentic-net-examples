using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RemoveOleObjectByLabel
{
    static void Main()
    {
        // Load the workbook containing OLE objects
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Define the label of the OLE object that should be removed
        string unwantedLabel = "UnwantedLabel";

        // Iterate through all worksheets (or target a specific one)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int indexToRemove = -1;

            // Search for the OLE object with the specified label
            for (int i = 0; i < sheet.OleObjects.Count; i++)
            {
                OleObject ole = sheet.OleObjects[i];
                if (ole.Label == unwantedLabel)
                {
                    indexToRemove = i;
                    break;
                }
            }

            // If found, remove it using RemoveAt
            if (indexToRemove >= 0)
            {
                sheet.OleObjects.RemoveAt(indexToRemove);
                Console.WriteLine($"Removed OLE object with label '{unwantedLabel}' from sheet '{sheet.Name}'.");
            }
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}