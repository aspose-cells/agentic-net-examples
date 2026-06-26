using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Add a custom Boolean property named "IsReviewed" with value true
        workbook.CustomDocumentProperties.Add("IsReviewed", true);

        // Save the workbook with the new property
        workbook.Save("output.xlsx");
    }
}