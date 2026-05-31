using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

public class Review
{
    public string Reviewer { get; set; }
    public string Comment { get; set; }
    public int Rating { get; set; }
}

public class SmartMarkerForeachDemo
{
    public static void Main()
    {
        // Create a template workbook with smart markers using foreach syntax
        Workbook template = new Workbook();
        Worksheet ws = template.Worksheets[0];
        Cells cells = ws.Cells;

        // Header row
        cells["A1"].PutValue("Reviewer");
        cells["B1"].PutValue("Comment");
        cells["C1"].PutValue("Rating");

        // Row that will be repeated for each item in the collection
        cells["A2"].PutValue("&=Reviews.Reviewer");
        cells["B2"].PutValue("&=Reviews.Comment");
        cells["C2"].PutValue("&=Reviews.Rating");

        // Save the template to a memory stream (lifecycle rule)
        MemoryStream templateStream = new MemoryStream();
        template.Save(templateStream, SaveFormat.Xlsx);
        templateStream.Position = 0;

        // Load the template into WorkbookDesigner
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = new Workbook(templateStream);

        // Prepare a variable‑length collection of product reviews
        List<Review> reviews = new List<Review>
        {
            new Review { Reviewer = "Alice",   Comment = "Excellent product!", Rating = 5 },
            new Review { Reviewer = "Bob",     Comment = "Very good, fast shipping.", Rating = 4 },
            new Review { Reviewer = "Charlie", Comment = "Average quality.", Rating = 3 }
        };

        // Bind the collection to the smart marker name "Reviews"
        designer.SetDataSource("Reviews", reviews);

        // Process the smart markers – the foreach syntax expands rows automatically
        designer.Process();

        // Save the final workbook (lifecycle rule)
        designer.Workbook.Save("ProductReviews.xlsx");
    }
}