using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsRelatedArticlesDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // create a new workbook instance

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            // Simple text placeholders
            sheet.Cells["A1"].PutValue("Title");
            sheet.Cells["A2"].PutValue("Author");
            sheet.Cells["A3"].PutValue("Date");
            sheet.Cells["B1"].PutValue("PlaceholderTitle");
            sheet.Cells["B2"].PutValue("PlaceholderAuthor");
            sheet.Cells["B3"].PutValue("PlaceholderDate");

            // Add a table that will be inserted later
            DataTable dt = new DataTable("Articles");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Title", typeof(string));
            dt.Columns.Add("Link", typeof(string));
            dt.Rows.Add(1, "Article One", "https://example.com/1");
            dt.Rows.Add(2, "Article Two", "https://example.com/2");
            dt.Rows.Add(3, "Article Three", "https://example.com/3");

            // Place a placeholder where the table will be inserted
            sheet.Cells["A5"].PutValue("ArticlesTable");

            // ---------- Perform replacements ----------
            // Replace simple text placeholders
            workbook.Replace("PlaceholderTitle", "Understanding Aspose.Cells");
            workbook.Replace("PlaceholderAuthor", "John Doe");
            workbook.Replace("PlaceholderDate", DateTime.Now.ToString("yyyy-MM-dd"));

            // Replace the placeholder with the DataTable (will expand starting at the placeholder cell)
            workbook.Replace("ArticlesTable", dt);

            // ---------- Save the workbook ----------
            // The Save method is the standard lifecycle operation for persisting the workbook
            workbook.Save("RelatedArticlesDemo.xlsx");
        }
    }
}