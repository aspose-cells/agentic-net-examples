using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerErrorHandling
{
    class Program
    {
        static void Main()
        {
            // Load a workbook that contains smart markers
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a sample data source (replace with actual data as needed)
            ArrayList data = new ArrayList();
            data.Add(new { Name = "John Doe", Age = 30 });
            data.Add(new { Name = "Jane Smith", Age = 28 });

            // Bind the data source to a name used in the smart markers
            designer.SetDataSource("Employees", data);

            try
            {
                // Process the smart markers; any syntax errors in the markers will throw an exception
                designer.Process();
                Console.WriteLine("Smart markers processed successfully.");
            }
            catch (CellsException ex) when (ex.Code == ExceptionType.InvalidData)
            {
                // Specific handling for malformed smart marker syntax
                Console.WriteLine("Smart marker syntax error detected:");
                Console.WriteLine($"Message: {ex.Message}");
                Console.WriteLine($"Error Code: {ex.Code}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine("An unexpected error occurred while processing smart markers:");
                Console.WriteLine($"Message: {ex.Message}");
            }

            // Save the workbook regardless of processing outcome
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved to output.xlsx");
        }
    }
}