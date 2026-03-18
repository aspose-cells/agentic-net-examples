using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomDateTimeProperty
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add a custom document property of type DateTime
            DateTime reportDate = DateTime.Now;
            DocumentProperty dateProp = workbook.CustomDocumentProperties.Add("ReportGenerated", reportDate);

            // Display information about the added property
            Console.WriteLine($"Property Name : {dateProp.Name}");
            Console.WriteLine($"Property Type : {dateProp.Type}");
            Console.WriteLine($"Property Value: {dateProp.ToDateTime():O}");

            // Save the workbook to an XLSX file (lifecycle: save)
            workbook.Save("CustomDateTimeProperty.xlsx", SaveFormat.Xlsx);
        }
    }
}