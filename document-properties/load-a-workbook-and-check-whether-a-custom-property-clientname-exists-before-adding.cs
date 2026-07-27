using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Check if the custom property "ClientName" already exists
        bool hasClientName = workbook.CustomDocumentProperties.Contains("ClientName");

        // Add the property only if it does not exist
        if (!hasClientName)
        {
            // Add a new custom document property of type string
            workbook.CustomDocumentProperties.Add("ClientName", "Acme Corp");
        }

        // Save the workbook (overwrites the original or saves to a new file)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}