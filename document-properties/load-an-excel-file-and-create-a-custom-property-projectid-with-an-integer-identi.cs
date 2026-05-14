using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Add a custom document property named "ProjectId" with an integer value
        workbook.CustomDocumentProperties.Add("ProjectId", 12345);

        // Save the workbook with the new property
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}