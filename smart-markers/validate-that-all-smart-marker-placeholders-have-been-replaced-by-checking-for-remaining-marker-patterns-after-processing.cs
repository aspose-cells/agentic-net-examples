using System;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerValidationDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a template workbook with smart markers in memory
            // ------------------------------------------------------------
            Workbook templateWorkbook = new Workbook();
            Worksheet sheet = templateWorkbook.Worksheets[0];
            // Smart markers using the old syntax (will be processed)
            sheet.Cells["A1"].PutValue("&=Employees.Name");
            sheet.Cells["B1"].PutValue("&=Employees.Age");

            // Save the template to a memory stream (simulating a file load)
            using (MemoryStream templateStream = new MemoryStream())
            {
                templateWorkbook.Save(templateStream, SaveFormat.Xlsx);
                templateStream.Position = 0; // Reset stream for reading

                // ------------------------------------------------------------
                // 2. Load the template into a WorkbookDesigner
                // ------------------------------------------------------------
                WorkbookDesigner designer = new WorkbookDesigner();
                designer.Workbook = new Workbook(templateStream);

                // ------------------------------------------------------------
                // 3. Prepare a data source matching the smart markers
                // ------------------------------------------------------------
                DataTable employeeTable = new DataTable("Employees");
                employeeTable.Columns.Add("Name", typeof(string));
                employeeTable.Columns.Add("Age", typeof(int));
                employeeTable.Rows.Add("John Doe", 30);
                employeeTable.Rows.Add("Jane Smith", 28);

                // Bind the data source to the designer
                designer.SetDataSource(employeeTable);

                // ------------------------------------------------------------
                // 4. Process the smart markers (replace placeholders)
                // ------------------------------------------------------------
                designer.Process();

                // ------------------------------------------------------------
                // 5. Validate that no smart marker placeholders remain
                // ------------------------------------------------------------
                string[] remainingMarkers = designer.GetSmartMarkers();

                if (remainingMarkers.Length == 0)
                {
                    Console.WriteLine("All smart markers have been successfully replaced.");
                }
                else
                {
                    Console.WriteLine("The following smart markers were not replaced:");
                    foreach (string marker in remainingMarkers)
                    {
                        Console.WriteLine(marker);
                    }
                }

                // ------------------------------------------------------------
                // 6. Save the processed workbook
                // ------------------------------------------------------------
                designer.Workbook.Save("ProcessedOutput.xlsx");
                Console.WriteLine("Processed workbook saved as 'ProcessedOutput.xlsx'.");
            }
        }
    }
}