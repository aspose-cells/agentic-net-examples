// Title: Import a variable‑length List of product reviews into an Excel template using foreach smart markers with Aspose.Cells in C#
// Common Searches: how to bind a List<Review> to Excel smart markers using foreach syntax in Aspose.Cells C# | populate an Excel template with a dynamic collection using WorkbookDesigner and foreach smart markers | convert a List of objects to a DataSet for Aspose.Cells smart markers processing in C# | use foreach smart markers to generate rows for product reviews in an Excel workbook | Aspose.Cells example for variable length data with foreach smart markers in C#
// Tags: foreach smart markers Aspose.Cells C# | WorkbookDesigner data source from List | List to DataTable conversion Aspose.Cells | populate Excel template with variable length collection | dynamic rows generation using smart markers

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel template that contains foreach smart markers, creates a List of Review objects, converts the list into a DataTable within a DataSet, sets this DataSet as the data source for WorkbookDesigner, processes the smart markers to fill the workbook with the review data, and saves the populated workbook to a new file.
public class Review
{
    public string Product { get; set; } = string.Empty;
    public int Rating { get; set; }
    public string Comment { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        try
        {
            const string templatePath = "Template.xlsx";
            const string resultPath = "Result.xlsx";

            // Verify that the template file exists to avoid FileNotFoundException
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"Template file not found: {templatePath}");
                return;
            }

            // Load the template workbook that contains smart markers with foreach syntax
            Workbook workbook = new Workbook(templatePath);

            // Prepare a variable‑length collection of product reviews
            List<Review> reviews = new List<Review>
            {
                new Review { Product = "Laptop", Rating = 5, Comment = "Excellent performance." },
                new Review { Product = "Headphones", Rating = 4, Comment = "Good sound quality." },
                new Review { Product = "Mouse", Rating = 3, Comment = "Average ergonomics." }
            };

            // Convert the list to a DataTable and add it to a DataSet (required by older WorkbookDesigner overloads)
            DataTable table = new DataTable("Reviews");
            table.Columns.Add("Product", typeof(string));
            table.Columns.Add("Rating", typeof(int));
            table.Columns.Add("Comment", typeof(string));

            foreach (var r in reviews)
            {
                table.Rows.Add(r.Product, r.Rating, r.Comment);
            }

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(table);

            // Process the smart markers using WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dataSet);
            designer.Process();

            // Save the populated workbook
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to {resultPath}");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
