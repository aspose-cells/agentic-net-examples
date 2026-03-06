using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Add various custom document properties (visible in Document Information Panel)
            workbook.CustomDocumentProperties.Add("ProjectName", "Alpha");
            workbook.CustomDocumentProperties.Add("Revision", 3);
            workbook.CustomDocumentProperties.Add("LastReviewed", DateTime.Now);
            workbook.CustomDocumentProperties.Add("IsApproved", true);
            workbook.CustomDocumentProperties.Add("Score", 4.75);

            // Save the workbook (save rule)
            workbook.Save("CustomPropertiesDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}