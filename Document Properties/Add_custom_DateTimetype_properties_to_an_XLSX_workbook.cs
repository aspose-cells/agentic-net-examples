using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsCustomDateTimeProperties
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add custom document properties of DateTime type
            // The Add(string, DateTime) overload creates a PropertyType.DateTime property
            DateTime reportGenerated = DateTime.Now;
            workbook.CustomDocumentProperties.Add("ReportGenerated", reportGenerated);

            DateTime reviewDate = new DateTime(2023, 12, 31, 15, 30, 0);
            workbook.CustomDocumentProperties.Add("ReviewDate", reviewDate);

            // Verify that the properties were added (optional)
            DocumentProperty prop1 = workbook.CustomDocumentProperties["ReportGenerated"];
            Console.WriteLine($"Name: {prop1.Name}, Type: {prop1.Type}, Value: {prop1.ToDateTime():O}");

            DocumentProperty prop2 = workbook.CustomDocumentProperties["ReviewDate"];
            Console.WriteLine($"Name: {prop2.Name}, Type: {prop2.Type}, Value: {prop2.ToDateTime():O}");

            // Save the workbook (the custom properties persist in the file)
            workbook.Save("CustomDateTimeProperties.xlsx");
        }
    }
}