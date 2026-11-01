// Title: Merging multiple JSON arrays into a single DataTable and binding it to Aspose.Cells smart markers with C#
// AI Prompts: Transform each JSON array into a DataTable, append the rows to a new table, and bind this table to the 'People' smart marker using WorkbookDesigner. | Create a helper method that converts a JSON string to a typed DataTable and reuse it to supply combined data for Aspose.Cells smart markers. | Load an Excel template, assign the combined DataTable to the smart marker data source, process the markers, and save the resulting workbook.
// Common Searches: How to combine multiple JSON arrays into one DataTable for Aspose.Cells smart markers in C# | Set a merged DataTable as a custom data source for WorkbookDesigner smart markers | Convert JSON array to DataTable and populate an Excel template with Aspose.Cells | Aspose.Cells smart markers using combined JSON data source
// Tags: aspocells json datatable merge | c# workbookdesigner setdatasource | excel smart markers data source json | json array to datatable aspocells | combined datatable smart markers c#

using System;
using System.Data;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

// The example loads an Excel template, converts two JSON arrays into DataTables, merges them into a single DataTable, sets this merged table as the 'People' custom data source for Aspose.Cells smart markers via WorkbookDesigner, processes the markers, and saves the populated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string templatePath = "Template.xlsx";
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the Excel template that contains smart markers.
            Workbook workbook = new Workbook(templatePath);

            // Example JSON arrays that need to be merged.
            string jsonArray1 = "[{\"Name\":\"John\",\"Age\":30},{\"Name\":\"Alice\",\"Age\":25}]";
            string jsonArray2 = "[{\"Name\":\"Bob\",\"Age\":28},{\"Name\":\"Eve\",\"Age\":22}]";

            // Convert JSON arrays to DataTables.
            DataTable table1 = JsonToDataTable(jsonArray1);
            DataTable table2 = JsonToDataTable(jsonArray2);

            // Merge tables.
            DataTable mergedTable = table1.Clone(); // copies column definitions
            foreach (DataRow row in table1.Rows) mergedTable.ImportRow(row);
            foreach (DataRow row in table2.Rows) mergedTable.ImportRow(row);

            // Set the merged DataTable as a custom data source for smart markers.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("People", mergedTable);
            designer.Process();

            // Save the populated workbook.
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Converts a JSON array of objects to a DataTable.
    private static DataTable JsonToDataTable(string json)
    {
        DataTable dt = new DataTable();

        try
        {
            using (JsonDocument doc = JsonDocument.Parse(json))
            {
                JsonElement root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Array || root.GetArrayLength() == 0) return dt;

                // Determine columns from the first element.
                foreach (JsonProperty prop in root[0].EnumerateObject())
                {
                    dt.Columns.Add(prop.Name, GetClrType(prop.Value));
                }

                // Populate rows.
                foreach (JsonElement element in root.EnumerateArray())
                {
                    DataRow row = dt.NewRow();
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (element.TryGetProperty(col.ColumnName, out JsonElement val))
                        {
                            row[col.ColumnName] = ConvertJsonValue(val, col.DataType);
                        }
                    }
                    dt.Rows.Add(row);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"JSON parsing error: {ex.Message}");
        }

        return dt;
    }

    // Maps JsonValueKind to a .NET type.
    private static Type GetClrType(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.Number => element.TryGetInt32(out _) ? typeof(int) : typeof(double),
            JsonValueKind.True or JsonValueKind.False => typeof(bool),
            _ => typeof(string)
        };
    }

    // Converts a JsonElement to the appropriate .NET value.
    private static object ConvertJsonValue(JsonElement element, Type targetType)
    {
        try
        {
            if (targetType == typeof(int) && element.TryGetInt32(out int i)) return i;
            if (targetType == typeof(double) && element.TryGetDouble(out double d)) return d;
            if (targetType == typeof(bool))
            {
                if (element.ValueKind == JsonValueKind.True) return true;
                if (element.ValueKind == JsonValueKind.False) return false;
            }
            return element.GetString() ?? string.Empty;
        }
        catch
        {
            return DBNull.Value;
        }
    }
}
