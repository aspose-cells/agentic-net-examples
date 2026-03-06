using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class RemoveCustomDocumentProperties
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Load the workbook (using the provided load rule)
        Workbook workbook = new Workbook(inputPath);

        // Get the collection of custom document properties
        CustomDocumentPropertyCollection customProps = workbook.CustomDocumentProperties;

        // Remove all custom document properties by iterating backwards
        for (int i = customProps.Count - 1; i >= 0; i--)
        {
            customProps.RemoveAt(i);
        }

        // Save the workbook (using the provided save rule)
        workbook.Save(outputPath);
    }
}