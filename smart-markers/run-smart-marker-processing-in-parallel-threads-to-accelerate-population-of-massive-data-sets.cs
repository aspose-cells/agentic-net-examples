// Title: Accelerate massive smart‑marker population in Excel using parallel threads with Aspose.Cells for C#
// AI Prompts: Generate C# code that uses Parallel.For to split a large DataTable into chunks, clone an Excel template per thread, assign each chunk to a WorkbookDesigner smart‑marker data source, process the markers, and save each result workbook with a unique filename. | Show how to create a 200,000‑row DataTable, configure a chunk size, and implement multithreaded smart‑marker processing with Aspose.Cells WorkbookDesigner in a .NET application. | Explain how to safely clone a Workbook, set the "Data" smart‑marker source, invoke Designer.Process, and write output files concurrently without race conditions.
// Common Searches: how to use Aspose.Cells WorkbookDesigner with Parallel.For for large data sets in C# | multithreaded smart marker processing in Excel using Aspose.Cells .NET | split DataTable into chunks for parallel smart marker population in Aspose.Cells | best practices for parallel Excel file generation with Aspose.Cells smart markers | performance tuning Aspose.Cells smart markers with massive DataTable
// Tags: parallel smart-marker processing Aspose.Cells | chunked DataTable handling for Excel templates | WorkbookDesigner multithreaded execution | massive data set Excel generation .NET | thread-safe workbook cloning Aspose.Cells

using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace SmartMarkerParallelProcessing
{
    // The example loads an Excel template, builds a 200,000‑row DataTable, divides it into configurable 5,000‑row chunks, and uses Parallel.For to clone the template per thread, assign each chunk as the "Data" smart‑marker source via WorkbookDesigner, process the markers, and save each chunk to a distinct Result_n.xlsx file, dramatically speeding up large‑scale smart‑marker population.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel template that contains smart markers
            string templatePath = "template.xlsx";

            // Directory where the generated workbooks will be saved
            string outputDirectory = "output";
            Directory.CreateDirectory(outputDirectory);

            // Load or generate a massive DataTable that will be used as the data source
            DataTable massiveData = GetMassiveDataTable();

            // Define the size of each data chunk that will be processed by a separate thread
            const int chunkSize = 5000; // adjust based on memory and performance testing
            int totalRows = massiveData.Rows.Count;
            int chunkCount = (totalRows + chunkSize - 1) / chunkSize;

            // Process each chunk in parallel
            Parallel.For(0, chunkCount, chunkIndex =>
            {
                // Clone the template workbook for this thread
                Workbook workbook = new Workbook(templatePath);

                // Create a WorkbookDesigner to work with smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Prepare a slice of the massive data for this chunk
                DataTable slice = massiveData.Clone(); // copy schema only
                int startRow = chunkIndex * chunkSize;
                int endRow = Math.Min(startRow + chunkSize, totalRows);
                for (int row = startRow; row < endRow; row++)
                {
                    slice.ImportRow(massiveData.Rows[row]);
                }

                // Assign the data source to the smart marker name used in the template
                designer.SetDataSource("Data", slice);

                // Process smart markers in the workbook
                designer.Process();

                // Save the processed workbook to a unique file
                string outputPath = Path.Combine(outputDirectory, $"Result_{chunkIndex}.xlsx");
                workbook.Save(outputPath);
            });

            Console.WriteLine("Parallel smart marker processing completed.");
        }

        /// <returns>A DataTable filled with sample data.</returns>
        private static DataTable GetMassiveDataTable()
        {
            DataTable table = new DataTable("Data");
            table.Columns.Add("ID", typeof(int));
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Quantity", typeof(double));
            table.Columns.Add("Price", typeof(decimal));
            table.Columns.Add("Date", typeof(DateTime));

            // Simulate a large data set (e.g., 200,000 rows)
            int totalRows = 200_000;
            Random rnd = new Random();

            for (int i = 1; i <= totalRows; i++)
            {
                DataRow row = table.NewRow();
                row["ID"] = i;
                row["Name"] = $"Product {i}";
                row["Quantity"] = rnd.NextDouble() * 100;
                row["Price"] = (decimal)(rnd.NextDouble() * 500);
                row["Date"] = DateTime.Today.AddDays(-rnd.Next(0, 365));
                table.Rows.Add(row);
            }

            return table;
        }
    }
}
