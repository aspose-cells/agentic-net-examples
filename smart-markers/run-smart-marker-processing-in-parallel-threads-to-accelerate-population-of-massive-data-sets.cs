using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

namespace SmartMarkerParallelProcessing
{
    // Sample data class used in smart markers
    public class Record
    {
        public string Category { get; set; }
        public double Amount { get; set; }

        public Record(string category, double amount)
        {
            Category = category;
            Amount = amount;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the template workbook that contains smart markers.
            // The template must define a named range "_CellsSmartMarkers"
            // covering the area where the smart markers will be expanded.
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Enable multi‑thread reading on the cells collection.
            // This allows the internal reading of cell values to be performed
            // concurrently when the smart marker engine accesses the worksheet.
            workbook.Worksheets[0].Cells.MultiThreadReading = true;

            // Prepare a massive data set (e.g., 1,000,000 rows).
            const int totalRows = 1_000_000;
            List<Record> data = new List<Record>(totalRows);
            Random rnd = new Random();
            string[] categories = { "Food", "Travel", "Utilities", "Entertainment", "Other" };
            for (int i = 0; i < totalRows; i++)
            {
                data.Add(new Record(
                    categories[rnd.Next(categories.Length)],
                    Math.Round(rnd.NextDouble() * 1000, 2)));
            }

            // Configure the designer.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // When LineByLine is false the designer works on the named range
                // "_CellsSmartMarkers". This is required for parallel processing.
                LineByLine = false
            };

            // Bind the massive data set to the smart marker name "Data".
            designer.SetDataSource("Data", data);

            // Process the smart markers.
            // Because MultiThreadReading is enabled, the engine can read cell
            // information in parallel, which speeds up the population of the
            // large data set.
            designer.Process();

            // Save the populated workbook.
            workbook.Save("PopulatedOutput.xlsx");
        }
    }
}