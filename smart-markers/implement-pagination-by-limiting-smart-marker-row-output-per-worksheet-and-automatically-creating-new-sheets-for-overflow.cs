// Title: How to paginate smart marker rows in Aspose.Cells .NET by limiting rows per worksheet and auto‑creating new sheets
// AI Prompts: Generate C# code that uses WorkbookDesigner to apply smart markers and split the data into multiple worksheets, capping each sheet at a specific row count. | Write a routine that copies a template worksheet for each data chunk and processes only that sheet with WorkbookDesigner to achieve pagination. | Create a helper that divides a collection into chunks and assigns each chunk to a separate worksheet using Aspose.Cells smart markers.
// Common Searches: Aspose.Cells paginate smart marker output across multiple worksheets | C# limit rows per sheet when using smart markers with WorkbookDesigner | how to automatically create new Excel sheets for overflow data in Aspose.Cells | split large data set into pages using smart markers in Aspose.Cells .NET | copy template sheet for each data chunk Aspose.Cells smart markers pagination
// Tags: smart marker pagination Aspose.Cells | limit rows per worksheet WorkbookDesigner | copy template worksheet for data overflow | process specific sheet with WorkbookDesigner | chunk list for Excel export Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaginationDemo
{
    // Sample data class used as a data source for smart markers
    // The example loads a template workbook containing a smart‑marker range, generates a list of employees, splits the list into chunks of up to 20 rows, copies the template sheet for each subsequent chunk, assigns each chunk as the data source for WorkbookDesigner, processes only the corresponding worksheet, and saves the paginated workbook as PaginatedResult.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    public class Program
    {
        // Maximum number of data rows that should appear on a single worksheet
        private const int MaxRowsPerSheet = 20;

        public static void Main()
        {
            // Load the template workbook that contains smart markers.
            // The template must have a named range "_CellsSmartMarkers" covering the row that will be repeated.
            Workbook workbook = new Workbook("template.xlsx");

            // Prepare a large list of employees to demonstrate pagination.
            List<Employee> allEmployees = GenerateSampleData(73); // e.g., 73 rows

            // Reference to the original template sheet (index 0)
            int templateSheetIndex = 0;

            // Split the data into chunks based on MaxRowsPerSheet
            List<List<Employee>> chunks = SplitIntoChunks(allEmployees, MaxRowsPerSheet);

            // Process each chunk on its own worksheet
            for (int i = 0; i < chunks.Count; i++)
            {
                int targetSheetIndex;

                if (i == 0)
                {
                    // First chunk uses the original template sheet
                    targetSheetIndex = templateSheetIndex;
                }
                else
                {
                    // Subsequent chunks: copy the template sheet to create a new sheet
                    targetSheetIndex = workbook.Worksheets.AddCopy(templateSheetIndex);
                    // Optionally rename the new sheet for clarity
                    workbook.Worksheets[targetSheetIndex].Name = $"Page_{i + 1}";
                }

                // Create a new designer for the current workbook
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // Use range smart markers (LineByLine = false) so that the named range is respected
                    LineByLine = false
                };

                // Set the data source for the current chunk.
                // The name "RootData" must match the smart marker prefix used in the template (e.g., &RootData.Name)
                designer.SetDataSource("RootData", chunks[i]);

                // Process only the target sheet. The second parameter (true) preserves unrecognized markers.
                designer.Process(targetSheetIndex, true);
            }

            // Save the paginated workbook.
            workbook.Save("PaginatedResult.xlsx");
        }

        // Generates a list of dummy employees for demonstration purposes.
        private static List<Employee> GenerateSampleData(int count)
        {
            var list = new List<Employee>();
            for (int i = 1; i <= count; i++)
            {
                list.Add(new Employee
                {
                    Name = $"Employee {i}",
                    Age = 20 + (i % 30),
                    Department = $"Dept {(i % 5) + 1}"
                });
            }
            return list;
        }

        // Splits a list into smaller lists each containing at most 'size' elements.
        private static List<List<Employee>> SplitIntoChunks(List<Employee> source, int size)
        {
            var chunks = new List<List<Employee>>();
            for (int i = 0; i < source.Count; i += size)
            {
                int chunkSize = Math.Min(size, source.Count - i);
                chunks.Add(source.GetRange(i, chunkSize));
            }
            return chunks;
        }
    }
}
