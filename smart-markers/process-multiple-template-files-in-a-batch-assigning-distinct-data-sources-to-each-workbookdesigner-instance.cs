// Title: Batch process multiple Excel templates with individual DataTable sources using Aspose.Cells WorkbookDesigner in C#
// Common Searches: C# Aspose.Cells generate reports from several Excel templates each with its own DataTable | How to assign different data sources to WorkbookDesigner instances in a loop | Batch fill smart markers in multiple .xlsx files using Aspose.Cells | Save processed Excel workbooks to a specific output folder with Aspose.Cells C# | Load and process multiple Excel template files with distinct data tables using Aspose.Cells
// Tags: batch processing Excel templates Aspose.Cells C# | WorkbookDesigner set distinct DataTable per template | smart markers multiple workbooks Aspose.Cells | save processed workbook to output directory C# | assign data source to Excel placeholder Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace BatchWorkbookDesigner
{
    // The example iterates over a collection of Excel template files, loads each into a WorkbookDesigner, retrieves a matching DataTable from a dictionary, assigns it to the "DataSource" smart marker, processes the template, and saves the resulting workbook with a "_Processed.xlsx" suffix to a designated output folder.
    class Program
    {
        static void Main(string[] args)
        {
            // Define template files and their corresponding data sources.
            var templateFiles = new List<string>
            {
                @"C:\Templates\Report1.xlsx",
                @"C:\Templates\Report2.xlsx",
                @"C:\Templates\Report3.xlsx"
            };

            // Example data sources: each template gets its own DataTable.
            var dataSources = new Dictionary<string, DataTable>
            {
                {
                    @"C:\Templates\Report1.xlsx",
                    CreateSampleDataTable("SalesData1")
                },
                {
                    @"C:\Templates\Report2.xlsx",
                    CreateSampleDataTable("SalesData2")
                },
                {
                    @"C:\Templates\Report3.xlsx",
                    CreateSampleDataTable("SalesData3")
                }
            };

            // Output directory for processed workbooks.
            string outputDir = @"C:\ProcessedReports";

            // Ensure the output directory exists.
            try
            {
                Directory.CreateDirectory(outputDir);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output directory '{outputDir}': {ex.Message}");
                return;
            }

            // Process each template with its distinct data source.
            foreach (var templatePath in templateFiles)
            {
                // Verify that the template file exists.
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    continue;
                }

                try
                {
                    // Load the template into a WorkbookDesigner instance.
                    var designer = new WorkbookDesigner
                    {
                        Workbook = new Workbook(templatePath) // Load template
                    };

                    // Retrieve the specific data source for this template.
                    if (!dataSources.TryGetValue(templatePath, out DataTable dataTable))
                    {
                        Console.WriteLine($"No data source found for template: {templatePath}");
                        continue;
                    }

                    // Assign the data source to the designer.
                    // The name "DataSource" must match the placeholder used in the template.
                    designer.SetDataSource("DataSource", dataTable);

                    // Process the template (merge data with placeholders).
                    designer.Process();

                    // Determine output file name.
                    string fileName = Path.GetFileNameWithoutExtension(templatePath);
                    string outputPath = Path.Combine(outputDir, $"{fileName}_Processed.xlsx");

                    // Save the processed workbook.
                    designer.Workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Processed workbook saved to: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                }
            }
        }

        // Helper method to create a sample DataTable.
        private static DataTable CreateSampleDataTable(string tableName)
        {
            var table = new DataTable(tableName);
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Quantity", typeof(int));
            table.Columns.Add("Price", typeof(double));

            // Add sample rows.
            table.Rows.Add("Apple", 120, 0.5);
            table.Rows.Add("Banana", 85, 0.3);
            table.Rows.Add("Cherry", 60, 1.2);

            return table;
        }
    }
}
