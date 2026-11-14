// Title: Create an Excel workbook from a nested Parent‑Child object hierarchy using Aspose.Cells smart marker range syntax in C#
// Common Searches: Aspose.Cells C# smart markers example with nested collections | How to export a parent‑child hierarchy to Excel using smart marker range syntax | C# WorkbookDesigner populate Excel from related DataSet tables | Map List<Parent> and List<Child> to Excel with Aspose.Cells smart markers | Generate Excel file from hierarchical objects using Aspose.Cells | Smart markers for hierarchical data in Aspose.Cells C# tutorial
// Tags: smart markers hierarchical data Aspose.Cells | WorkbookDesigner populate Excel from DataSet C# | nested collections export to Excel Aspose.Cells | parent child DataRelation Excel generation | smart marker range syntax C# example

using System;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace SmartMarkerDemo
{
    // Define the data model
    // The sample creates a new Workbook, adds column headers, and inserts smart markers (&-Parents.Name, &-Children.Name, &-Children.Age) to iterate over a Parent collection and its nested Child collection. It builds a DataSet with Parent and Child tables linked by a DataRelation, assigns the DataSet to a WorkbookDesigner, processes the smart markers to fill the worksheet, and saves the result as NestedDataSmartMarkers.xlsx.
    public class Parent
    {
        public string? Name { get; set; }
        public List<Child>? Children { get; set; }
    }

    public class Child
    {
        public string? Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Write headers
                sheet.Cells["A1"].PutValue("Parent Name");
                sheet.Cells["B1"].PutValue("Child Name");
                sheet.Cells["C1"].PutValue("Child Age");

                // 4. Insert smart markers.
                //    &-Parents.Name   -> iterates over the Parent collection
                //    &-Children.Name  -> iterates over the Child collection of the current Parent
                //    &-Children.Age   -> iterates over the Child collection of the current Parent
                sheet.Cells["A2"].PutValue("&-Parents.Name");
                sheet.Cells["B2"].PutValue("&-Children.Name");
                sheet.Cells["C2"].PutValue("&-Children.Age");

                // 5. Prepare sample data
                List<Parent> parents = new List<Parent>
                {
                    new Parent
                    {
                        Name = "Parent A",
                        Children = new List<Child>
                        {
                            new Child { Name = "Child A1", Age = 10 },
                            new Child { Name = "Child A2", Age = 12 }
                        }
                    },
                    new Parent
                    {
                        Name = "Parent B",
                        Children = new List<Child>
                        {
                            new Child { Name = "Child B1", Age = 8 },
                            new Child { Name = "Child B2", Age = 9 },
                            new Child { Name = "Child B3", Age = 11 }
                        }
                    }
                };

                // 6. Convert the hierarchical list to a DataSet for smart markers
                DataSet dataSet = BuildDataSet(parents);

                // 7. Use WorkbookDesigner to process smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(dataSet);
                designer.Process(); // Fills the smart markers with data

                // 8. Save the populated workbook
                string outputPath = "NestedDataSmartMarkers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to build a DataSet with Parent and Child tables and a relation
        private static DataSet BuildDataSet(List<Parent> parents)
        {
            DataSet ds = new DataSet();

            // Parent table
            DataTable parentTable = new DataTable("Parents");
            parentTable.Columns.Add("ParentID", typeof(int));
            parentTable.Columns.Add("Name", typeof(string));
            parentTable.PrimaryKey = new[] { parentTable.Columns["ParentID"] };

            // Child table
            DataTable childTable = new DataTable("Children");
            childTable.Columns.Add("ChildID", typeof(int));
            childTable.Columns.Add("ParentID", typeof(int));
            childTable.Columns.Add("Name", typeof(string));
            childTable.Columns.Add("Age", typeof(int));
            childTable.PrimaryKey = new[] { childTable.Columns["ChildID"] };

            int parentId = 1;
            int childId = 1;

            foreach (var parent in parents)
            {
                // Add parent row
                DataRow pRow = parentTable.NewRow();
                pRow["ParentID"] = parentId;
                pRow["Name"] = parent.Name ?? string.Empty;
                parentTable.Rows.Add(pRow);

                // Add child rows
                if (parent.Children != null)
                {
                    foreach (var child in parent.Children)
                    {
                        DataRow cRow = childTable.NewRow();
                        cRow["ChildID"] = childId;
                        cRow["ParentID"] = parentId;
                        cRow["Name"] = child.Name ?? string.Empty;
                        cRow["Age"] = child.Age;
                        childTable.Rows.Add(cRow);
                        childId++;
                    }
                }

                parentId++;
            }

            ds.Tables.Add(parentTable);
            ds.Tables.Add(childTable);

            // Define relation between Parents and Children
            DataRelation relation = new DataRelation(
                "Parent_Children",
                parentTable.Columns["ParentID"],
                childTable.Columns["ParentID"]);
            ds.Relations.Add(relation);

            return ds;
        }
    }
}
