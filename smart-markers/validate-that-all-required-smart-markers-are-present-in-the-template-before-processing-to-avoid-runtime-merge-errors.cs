// Title: C# code to validate all smart markers in an Excel template before processing with Aspose.Cells WorkbookDesigner
// Common Searches: how to check for missing smart marker fields in an Aspose.Cells Excel template using C# | C# validate smart markers against a DataTable data source before calling WorkbookDesigner | prevent runtime merge errors by verifying smart markers in a .NET Excel workbook | extract smart markers from an .xlsx file with regex and Aspose.Cells | debug smart marker data source mismatches in an Aspose.Cells C# application
// Tags: Aspose.Cells WorkbookDesigner data source verification | C# regex extraction of Excel smart markers | prevent runtime merge errors Aspose.Cells | smart marker field existence verification | Excel template smart marker consistency check

using System;
using System.Data;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerValidator
{
    // The sample loads an Excel template, registers a DataTable as a data source, scans the workbook with a regular expression to locate smart markers, and verifies that each marker's referenced data source and field exist. If any validation fails, processing stops to avoid runtime merge errors; otherwise the WorkbookDesigner processes the markers and saves the result.
    class Program
    {
        // Represents a parsed smart marker.
        private class SmartMarkerInfo
        {
            public string DataSourceName { get; set; }
            public string FieldName { get; set; }
            public string CellAddress { get; set; }
        }

        static void Main(string[] args)
        {
            // Expect at least two arguments: template file and output file.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: SmartMarkerValidator <template.xlsx> <output.xlsx>");
                return;
            }

            string templatePath = args[0];
            string outputPath = args[1];

            // Verify that the template file exists.
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(templatePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load template workbook: {ex.Message}");
                return;
            }

            // Create a WorkbookDesigner instance.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Keep a local map of data sources for validation.
            var dataSources = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

            // ------------------------------------------------------------
            // STEP 1: Prepare data sources.
            // ------------------------------------------------------------
            DataTable employeesTable = new DataTable("Employees");
            employeesTable.Columns.Add("Name", typeof(string));
            employeesTable.Columns.Add("Age", typeof(int));
            employeesTable.Columns.Add("Department", typeof(string));

            employeesTable.Rows.Add("John Doe", 30, "HR");
            employeesTable.Rows.Add("Jane Smith", 28, "Finance");
            employeesTable.Rows.Add("Bob Johnson", 45, "IT");

            // Register the data source with the designer and store it locally.
            designer.SetDataSource("Employees", employeesTable);
            dataSources["Employees"] = employeesTable;

            // ------------------------------------------------------------
            // STEP 2: Scan the workbook for smart markers.
            // ------------------------------------------------------------
            List<SmartMarkerInfo> markers = ExtractSmartMarkers(workbook);

            if (markers.Count == 0)
            {
                Console.WriteLine("No smart markers were found in the template. Nothing to process.");
                return;
            }

            // ------------------------------------------------------------
            // STEP 3: Validate that each smart marker has a matching field.
            // ------------------------------------------------------------
            bool validationFailed = false;
            foreach (var marker in markers)
            {
                // Retrieve the data source object by name from the local map.
                if (!dataSources.TryGetValue(marker.DataSourceName, out object ds) || ds == null)
                {
                    Console.WriteLine($"[Error] Data source \"{marker.DataSourceName}\" referenced in cell {marker.CellAddress} does not exist.");
                    validationFailed = true;
                    continue;
                }

                // Currently we only support DataTable as a data source.
                if (ds is DataTable dt)
                {
                    if (!dt.Columns.Contains(marker.FieldName))
                    {
                        Console.WriteLine($"[Error] Field \"{marker.FieldName}\" not found in data source \"{marker.DataSourceName}\" (cell {marker.CellAddress}).");
                        validationFailed = true;
                    }
                }
                else
                {
                    Console.WriteLine($"[Warning] Data source \"{marker.DataSourceName}\" is of type {ds.GetType().Name}. Field validation is not implemented for this type (cell {marker.CellAddress}).");
                }
            }

            if (validationFailed)
            {
                Console.WriteLine("Smart marker validation failed. Aborting processing to avoid runtime merge errors.");
                return;
            }

            // ------------------------------------------------------------
            // STEP 4: Process the smart markers now that validation passed.
            // ------------------------------------------------------------
            try
            {
                designer.Process();
                workbook.Save(outputPath);
                Console.WriteLine($"Smart markers processed successfully. Output saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during processing: {ex.Message}");
            }
        }

        /// <param name="workbook">The workbook to scan.</param>
        /// <returns>A list of discovered smart markers.</returns>
        private static List<SmartMarkerInfo> ExtractSmartMarkers(Workbook workbook)
        {
            var markers = new List<SmartMarkerInfo>();
            // Regex to capture DataSource and Field parts, e.g., "&=Employees.Name"
            Regex smRegex = new Regex(@"&=\s*(\w+)\.(\w+)", RegexOptions.IgnoreCase);

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Iterate through all used cells only for performance.
                foreach (Cell cell in cells)
                {
                    if (cell.Type != CellValueType.IsString)
                        continue;

                    string cellValue = cell.StringValue;
                    if (string.IsNullOrEmpty(cellValue))
                        continue;

                    Match match = smRegex.Match(cellValue);
                    if (match.Success && match.Groups.Count == 3)
                    {
                        var info = new SmartMarkerInfo
                        {
                            DataSourceName = match.Groups[1].Value,
                            FieldName = match.Groups[2].Value,
                            CellAddress = $"{sheet.Name}!{cell.Name}"
                        };
                        markers.Add(info);
                    }
                }
            }

            return markers;
        }
    }
}
