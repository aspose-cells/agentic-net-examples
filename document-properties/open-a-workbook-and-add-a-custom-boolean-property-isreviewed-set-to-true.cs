using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add a custom Boolean property named "IsReviewed" with value true
        workbook.CustomDocumentProperties.Add("IsReviewed", true);

        // Save the workbook to a file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}