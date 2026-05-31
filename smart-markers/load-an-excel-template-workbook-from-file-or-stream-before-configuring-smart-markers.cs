using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerDemo
{
    // Sample data class used as a smart marker data source
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load template workbook from a file
            // -----------------------------------------------------------------
            string templatePath = "Template.xlsx"; // Path to the Excel template containing smart markers
            Workbook workbookFromFile = new Workbook(templatePath); // LoadOptions not required for default loading

            // -----------------------------------------------------------------
            // 2. Load template workbook from a memory stream
            // -----------------------------------------------------------------
            // First, create a temporary workbook and save it to a MemoryStream
            Workbook tempWorkbook = new Workbook();
            Worksheet tempSheet = tempWorkbook.Worksheets[0];
            tempSheet.Cells["A1"].PutValue("&Employee.Name");        // Smart marker for Name
            tempSheet.Cells["A2"].PutValue("&Employee.Age");         // Smart marker for Age
            tempSheet.Cells["A3"].PutValue("&Employee.Department"); // Smart marker for Department

            // Save the temporary workbook to a stream (Excel97-2003 format as required by SaveToStream)
            MemoryStream stream = tempWorkbook.SaveToStream();

            // Reset stream position before reading
            stream.Position = 0;

            // Load the workbook from the stream
            Workbook workbookFromStream = new Workbook(stream);

            // -----------------------------------------------------------------
            // 3. Prepare data source for smart markers
            // -----------------------------------------------------------------
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "Marketing" }
            };

            // -----------------------------------------------------------------
            // 4. Process smart markers using WorkbookDesigner (file‑based workbook)
            // -----------------------------------------------------------------
            WorkbookDesigner designerFile = new WorkbookDesigner();
            designerFile.Workbook = workbookFromFile;               // Assign loaded workbook
            designerFile.SetDataSource("Employee", employees);     // Bind data source
            designerFile.Process();                                // Populate smart markers

            // Save the processed workbook
            string outputFilePath = "ProcessedFromFile.xlsx";
            designerFile.Workbook.Save(outputFilePath, SaveFormat.Xlsx);

            // -----------------------------------------------------------------
            // 5. Process smart markers using WorkbookDesigner (stream‑based workbook)
            // -----------------------------------------------------------------
            WorkbookDesigner designerStream = new WorkbookDesigner();
            designerStream.Workbook = workbookFromStream;          // Assign loaded workbook
            designerStream.SetDataSource("Employee", employees);   // Bind same data source
            designerStream.Process();                              // Populate smart markers

            // Save the processed workbook
            string outputStreamPath = "ProcessedFromStream.xlsx";
            designerStream.Workbook.Save(outputStreamPath, SaveFormat.Xlsx);

            // Clean up
            stream.Dispose();

            Console.WriteLine("Processing completed.");
            Console.WriteLine($"File‑based output saved to: {outputFilePath}");
            Console.WriteLine($"Stream‑based output saved to: {outputStreamPath}");
        }
    }
}