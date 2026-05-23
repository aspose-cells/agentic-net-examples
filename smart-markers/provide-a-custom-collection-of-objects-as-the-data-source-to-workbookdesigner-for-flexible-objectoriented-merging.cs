using System;
using System.Collections.Generic;
using Aspose.Cells;

// Custom class representing a row of data
public class Product
{
    public string Name { get; set; }
    public double Price { get; set; }
    public int Quantity { get; set; }

    public Product(string name, double price, int quantity)
    {
        Name = name;
        Price = price;
        Quantity = quantity;
    }
}

public class CustomCollectionDemo
{
    public static void Run()
    {
        // 1. Create a WorkbookDesigner instance
        WorkbookDesigner designer = new WorkbookDesigner();

        // 2. Create a new workbook (template) and assign it to the designer
        designer.Workbook = new Workbook();

        // 3. Access the first worksheet
        Worksheet sheet = designer.Workbook.Worksheets[0];

        // 4. Define column headers
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["C1"].PutValue("Quantity");

        // 5. Place smart markers that reference the custom collection "Products"
        sheet.Cells["A2"].PutValue("&=Products.Name");
        sheet.Cells["B2"].PutValue("&=Products.Price");
        sheet.Cells["C2"].PutValue("&=Products.Quantity");

        // 6. Prepare a custom collection of Product objects
        List<Product> products = new List<Product>
        {
            new Product("Apple", 1.20, 50),
            new Product("Banana", 0.80, 120),
            new Product("Orange", 1.50, 75)
        };

        // 7. Bind the collection to the smart marker name "Products"
        designer.SetDataSource("Products", products);

        // 8. Process the smart markers to populate the worksheet
        designer.Process();

        // 9. Save the populated workbook
        designer.Workbook.Save("CustomCollectionOutput.xlsx");
    }
}

// Entry point for demonstration
class Program
{
    static void Main()
    {
        CustomCollectionDemo.Run();
    }
}