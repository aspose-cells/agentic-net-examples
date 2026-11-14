// Title: Step‑by‑step guide to create a workbook, populate rows and columns, insert a column chart, and save/load an XLSX file using Aspose.Cells for .NET (C#)
// Common Searches: asp.net c# generate XLSX workbook, add data to cells and insert column chart with Aspose.Cells | using Aspose.Cells .NET to save a workbook to a folder and then reload it | read a specific cell value after loading an Excel file with Aspose.Cells C# example | populate month and sales data in Aspose.Cells and create a column chart programmatically | Aspose.Cells chart NSeries Add range syntax C#
// Tags: Aspose.Cells workbook creation C# | cell data population Aspose.Cells | column chart generation Aspose.Cells | XLSX file saving Aspose.Cells | workbook loading and cell read Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

namespace AsposeCellsGuide
{
    // This example demonstrates how to create a new Workbook, fill the first worksheet with month and sales data, add a column chart based on that range, save the workbook as Output.xlsx, reload the file, and verify the content by reading cell A1 using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Step 1: Create a new workbook
                Workbook workbook = CreateWorkbook();

                // Step 2: Add sample data to the first worksheet
                PopulateData(workbook);

                // Step 3: Insert a chart based on the data
                AddChart(workbook);

                // Step 4: Save the workbook to a file
                string outputPath = "Output.xlsx";
                SaveWorkbook(workbook, outputPath);

                // Step 5: Load the workbook back from the file
                Workbook loaded = LoadWorkbook(outputPath);

                // Step 6: Verify that data was loaded correctly
                Console.WriteLine("Value in A1: " + loaded.Worksheets[0].Cells["A1"].StringValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Creates a new workbook with a default worksheet
        static Workbook CreateWorkbook()
        {
            return new Workbook();
        }

        // Populates sample data into the first worksheet
        static void PopulateData(Workbook wb)
        {
            Worksheet sheet = wb.Worksheets[0];

            // Header row
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            // Sample rows
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 1200, 1500, 1800, 1600, 2000 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
            }
        }

        // Adds a simple column chart based on the populated data
        static void AddChart(Workbook wb)
        {
            try
            {
                Worksheet sheet = wb.Worksheets[0];

                // Add a column chart positioned from row 7, column 0 to row 20, column 10
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Monthly Sales";

                // Define the data range for the series (values)
                chart.NSeries.Add("B2:B6", true);

                // Optional: set series name
                chart.NSeries[0].Name = "Sales";

                // Note: CategoryData property may not be available in some versions.
                // If needed, uncomment the line below and ensure the Aspose.Cells version supports it.
                // chart.NSeries[0].CategoryData = "A2:A6";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to add chart: {ex.Message}");
                throw;
            }
        }

        // Saves the workbook to the specified file path in XLSX format
        static void SaveWorkbook(Workbook wb, string filePath)
        {
            try
            {
                string directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                wb.Save(filePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
                throw;
            }
        }

        // Loads a workbook from the specified file path
        static Workbook LoadWorkbook(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The file '{filePath}' does not exist.");
            }

            try
            {
                return new Workbook(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                throw;
            }
        }
    }
}
