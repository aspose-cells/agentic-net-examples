// Title: Create an Excel report of product reviews with variable‑length rows using foreach smart markers in Aspose.Cells for .NET
// AI Prompts: Create an Excel workbook that repeats a template row for every Review item by using the '&=' marker syntax. | Assign a List<Review> to the smart marker name "Reviews" and set WorkbookDesigner.LineByLine to false to enable range smart markers. | Run the smart marker processing and write the populated workbook to a .xlsx file.
// Common Searches: Aspose.Cells foreach smart marker syntax for repeating rows based on a List<T> | How to bind a collection to a smart marker range in Aspose.Cells .NET | Export product review data to Excel using range smart markers with LineByLine false | Variable length data export example with Aspose.Cells smart markers
// Tags: range smart markers with variable length data | WorkbookDesigner LineByLine false | bind List<T> to smart marker Aspose.Cells | export product reviews to Excel | smart marker row repetition Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example builds a template workbook, defines a foreach smart‑marker row for a collection named Reviews, binds a List<Review> to that marker, configures WorkbookDesigner for range smart markers (LineByLine = false), processes the markers, and saves the result as ProductReviews.xlsx.
public class Review
{
    public string Reviewer { get; set; }
    public string Comment { get; set; }
    public int Rating { get; set; }
}

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a template workbook with smart markers
        // -------------------------------------------------
        Workbook template = new Workbook();
        Worksheet ws = template.Worksheets[0];
        Cells cells = ws.Cells;

        // Title
        cells["A1"].PutValue("Product Reviews");

        // Column headers (row 3)
        cells["A3"].PutValue("Reviewer");
        cells["B3"].PutValue("Comment");
        cells["C3"].PutValue("Rating");

        // Smart marker row (row 4) – foreach syntax (&=) repeats this row for each item in the collection
        cells["A4"].PutValue("&=Reviews.Reviewer");
        cells["B4"].PutValue("&=Reviews.Comment");
        cells["C4"].PutValue("&=Reviews.Rating");

        // Define the range that contains the smart markers.
        // When LineByLine = false the designer looks for a range named "_CellsSmartMarkers".
        ws.Cells.CreateRange("A4:C4").Name = "_CellsSmartMarkers";

        // -------------------------------------------------
        // 2. Set up the WorkbookDesigner
        // -------------------------------------------------
        WorkbookDesigner designer = new WorkbookDesigner
        {
            Workbook = template,
            // Use range smart markers (required for foreach syntax with variable‑length data)
            LineByLine = false
        };

        // -------------------------------------------------
        // 3. Prepare a variable‑length collection of reviews
        // -------------------------------------------------
        List<Review> reviews = new List<Review>
        {
            new Review { Reviewer = "Alice",   Comment = "Great product!", Rating = 5 },
            new Review { Reviewer = "Bob",     Comment = "Good value.",    Rating = 4 },
            new Review { Reviewer = "Charlie", Comment = "Average.",       Rating = 3 }
        };

        // Bind the collection to the smart marker name "Reviews"
        designer.SetDataSource("Reviews", reviews);

        // -------------------------------------------------
        // 4. Process the smart markers and save the result
        // -------------------------------------------------
        designer.Process();
        designer.Workbook.Save("ProductReviews.xlsx");
    }
}
