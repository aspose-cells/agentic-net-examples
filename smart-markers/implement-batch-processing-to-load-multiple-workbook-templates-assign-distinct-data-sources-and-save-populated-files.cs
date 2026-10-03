// Title: Generate multiple Excel reports from different templates with individual DataTable sources using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over a list of Excel template files, loads each with Aspose.Cells, inserts the matching DataTable into the first worksheet, and saves the workbook with a filename that combines the template name and the table name. | Refactor the batch example to employ Aspose.Cells smart markers so the DataTable is mapped automatically instead of using manual cell loops. | Add robust error handling to the batch processor: skip missing template files, log each step, and continue processing remaining items after any exception.
// Common Searches: c# Aspose.Cells batch generate reports from multiple templates with separate data tables | how to use Aspose.Cells to fill different Excel templates with DataTable in a loop | save Aspose.Cells workbook with dynamic filename based on template and data source | Aspose.Cells smart markers batch processing multiple workbooks .NET | error handling for missing Excel template files Aspose.Cells batch
// Tags: batch workbook generation Aspose.Cells | populate Excel template from DataTable Aspose.Cells | dynamic output filename Aspose.Cells | smart markers data insertion Aspose.Cells | error handling missing template Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookProcessor
{
    // The program loops through a collection of Excel template paths, loads each workbook with Aspose.Cells, clears the first worksheet, writes the associated DataTable (including column headers) into the sheet, and saves the result to an output folder using a filename that merges the template name and the DataTable name.
    class Program
    {
        static void Main(string[] args)
        {
            // Define template files and corresponding data sources
            var templates = new List<string>
            {
                @"C:\Templates\ReportTemplate1.xlsx",
                @"C:\Templates\ReportTemplate2.xlsx",
                @"C:\Templates\ReportTemplate3.xlsx"
            };

            // Example data sources – in real scenarios these could come from a database, service, etc.
            var dataSources = new List<DataTable>
            {
                CreateSampleDataTable("Sales Q1"),
                CreateSampleDataTable("Sales Q2"),
                CreateSampleDataTable("Sales Q3")
            };

            // Output folder
            string outputFolder = @"C:\GeneratedReports\";

            // Ensure output directory exists
            try
            {
                if (!Directory.Exists(outputFolder))
                {
                    Directory.CreateDirectory(outputFolder);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create output directory: {ex.Message}");
                return;
            }

            // Process each template with its data source
            for (int i = 0; i < templates.Count; i++)
            {
                string templatePath = templates[i];
                DataTable data = dataSources[i];

                // Verify template file exists
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template not found: {templatePath}");
                    continue;
                }

                try
                {
                    // Load the workbook template
                    Workbook workbook = new Workbook(templatePath);

                    // Ensure there is at least one worksheet
                    if (workbook.Worksheets.Count == 0)
                    {
                        Console.WriteLine($"No worksheets found in template: {templatePath}");
                        continue;
                    }

                    // Assume the first worksheet is the target for data import
                    Worksheet sheet = workbook.Worksheets[0];

                    // Clear existing data (optional, depending on template design)
                    sheet.Cells.Clear();

                    // Write DataTable to worksheet manually (headers + rows)
                    try
                    {
                        int rowIndex = 0;

                        // Write column headers
                        for (int col = 0; col < data.Columns.Count; col++)
                        {
                            sheet.Cells[rowIndex, col].PutValue(data.Columns[col].ColumnName);
                        }
                        rowIndex++;

                        // Write each data row
                        foreach (DataRow dr in data.Rows)
                        {
                            for (int col = 0; col < data.Columns.Count; col++)
                            {
                                sheet.Cells[rowIndex, col].PutValue(dr[col]);
                            }
                            rowIndex++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to write data into worksheet: {ex.Message}");
                        continue;
                    }

                    // Build the output file name (e.g., ReportTemplate1_Sales Q1.xlsx)
                    string templateFileName = Path.GetFileNameWithoutExtension(templatePath);
                    string outputFileName = $"{templateFileName}_{data.TableName}.xlsx";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the populated workbook
                    workbook.Save(outputPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Generated: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing template '{templatePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }

        // Helper method to create a sample DataTable for demonstration purposes
        private static DataTable CreateSampleDataTable(string tableName)
        {
            DataTable dt = new DataTable(tableName);
            dt.Columns.Add("Product", typeof(string));
            dt.Columns.Add("Quantity", typeof(int));
            dt.Columns.Add("Price", typeof(decimal));

            dt.Rows.Add("Widget A", 120, 9.99m);
            dt.Rows.Add("Widget B", 85, 14.50m);
            dt.Rows.Add("Widget C", 60, 7.25m);

            return dt;
        }
    }
}
