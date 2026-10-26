// Title: Set WorkbookDesigner.LineByLine = false to merge nested Parent‑Child objects as grouped records with Aspose.Cells for .NET
// Common Searches: how to disable linebyline in WorkbookDesigner for hierarchical data in Aspose.Cells C# | Aspose.Cells smart markers processing nested collections as grouped records | C# example merging parent child DataSet into Excel using WorkbookDesigner | using WorkbookDesigner to handle parent‑child relations without line‑by‑line output | Aspose.Cells template with smart markers for nested objects
// Tags: WorkbookDesigner LineByLine false | Aspose.Cells hierarchical smart markers | C# nested DataSet relation Excel export | grouped records processing Aspose.Cells | parent child smart markers Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // Sample data classes representing nested objects
    // The example demonstrates creating a WorkbookDesigner, setting LineByLine to false so that nested Parent‑Child collections are treated as grouped records, building a DataSet with Parent and Child tables linked by a DataRelation, assigning the DataSet as the data source, processing smart markers in the Excel template, and saving the resulting workbook.
    public class Parent
    {
        public string? Name { get; set; }
        public List<Child>? Children { get; set; }
    }

    public class Child
    {
        public string? Item { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Ensure the template file exists; create a minimal workbook if missing
                Workbook workbook;
                if (File.Exists(templatePath))
                {
                    workbook = new Workbook(templatePath);
                }
                else
                {
                    workbook = new Workbook();
                    // Optionally add a worksheet with markers for demonstration
                    Worksheet ws = workbook.Worksheets[0];
                    ws.Cells["A1"].PutValue("&=Parent.Name");
                    ws.Cells["A2"].PutValue("&=Child.Item");
                    ws.Cells["B2"].PutValue("&=Child.Quantity");
                }

                // Create a WorkbookDesigner instance for the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook)
                {
                    // Treat nested collections as grouped records
                    LineByLine = false
                };

                // Prepare sample nested data
                var data = new List<Parent>
                {
                    new Parent
                    {
                        Name = "Order A",
                        Children = new List<Child>
                        {
                            new Child { Item = "Apple", Quantity = 10 },
                            new Child { Item = "Banana", Quantity = 5 }
                        }
                    },
                    new Parent
                    {
                        Name = "Order B",
                        Children = new List<Child>
                        {
                            new Child { Item = "Orange", Quantity = 8 },
                            new Child { Item = "Grape", Quantity = 12 }
                        }
                    }
                };

                // Convert the nested list into a DataSet compatible with WorkbookDesigner
                DataSet ds = BuildDataSet(data);

                // Set the data source for the designer
                designer.SetDataSource(ds);

                // Process the markers in the workbook using the provided data
                designer.Process();

                // Save the resulting workbook
                workbook.Save(resultPath);
                Console.WriteLine($"Workbook saved to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Builds a DataSet with Parent and Child tables and a relation between them
        private static DataSet BuildDataSet(List<Parent> parents)
        {
            DataSet ds = new DataSet();

            // Parent table
            DataTable parentTable = new DataTable("Parent");
            parentTable.Columns.Add("Name", typeof(string));
            ds.Tables.Add(parentTable);

            // Child table
            DataTable childTable = new DataTable("Child");
            childTable.Columns.Add("Item", typeof(string));
            childTable.Columns.Add("Quantity", typeof(int));
            childTable.Columns.Add("ParentName", typeof(string)); // foreign key column
            ds.Tables.Add(childTable);

            // Populate tables
            foreach (var p in parents)
            {
                // Add parent row
                DataRow parentRow = parentTable.NewRow();
                parentRow["Name"] = p.Name ?? string.Empty;
                parentTable.Rows.Add(parentRow);

                // Add child rows linked to this parent
                if (p.Children != null)
                {
                    foreach (var c in p.Children)
                    {
                        DataRow childRow = childTable.NewRow();
                        childRow["Item"] = c.Item ?? string.Empty;
                        childRow["Quantity"] = c.Quantity;
                        childRow["ParentName"] = p.Name ?? string.Empty;
                        childTable.Rows.Add(childRow);
                    }
                }
            }

            // Create relation between Parent.Name and Child.ParentName
            DataRelation relation = new DataRelation(
                "Parent_Child",
                parentTable.Columns["Name"],
                childTable.Columns["ParentName"]);
            ds.Relations.Add(relation);

            return ds;
        }
    }
}
