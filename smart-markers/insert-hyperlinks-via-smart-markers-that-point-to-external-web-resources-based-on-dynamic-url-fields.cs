// Title: Insert external hyperlinks into Excel using Aspose.Cells smart markers with a DataTable in C#
// AI Prompts: Write C# code that uses Aspose.Cells WorkbookDesigner to replace a smart marker with clickable URLs sourced from a DataTable column. | Show how to place the '&="Hyperlink"' smart marker in a worksheet and bind it to a DataTable so that Excel cells become dynamic hyperlinks.
// Common Searches: Aspose.Cells C# bind DataTable to smart marker for hyperlink generation | Create clickable URLs in Excel with WorkbookDesigner and smart markers .NET | Replace smart marker with external link from DataTable using Aspose.Cells | C# example of inserting multiple hyperlinks into an Excel file via smart markers
// Tags: smart-marker hyperlink Aspose.Cells | WorkbookDesigner populate URLs from DataTable | generate external links Excel C# Aspose.Cells | dynamic hyperlink insertion Excel .NET

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsHyperlinkExample
{
    // The example creates a new workbook, adds a smart marker '&="Hyperlink"' to cell A1, builds a DataTable containing a 'Url' column with target addresses, assigns the table to a WorkbookDesigner as the data source named 'Links', processes the smart markers to replace the placeholder with actual clickable hyperlinks, and saves the workbook as 'HyperlinksOutput.xlsx'.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                var workbook = new Workbook();

                // Access the first worksheet.
                var worksheet = workbook.Worksheets[0];

                // Insert a smart marker that will be replaced by a hyperlink.
                // The syntax &="Hyperlink" tells Aspose.Cells to treat the cell as a hyperlink.
                // The column name "Url" (provided in the data source) will be used as the hyperlink address.
                worksheet.Cells["A1"].PutValue("&=\"Hyperlink\"");

                // Prepare the data source – a DataTable with a single column "Url" containing the target URLs.
                var dataTable = new DataTable();
                dataTable.Columns.Add("Url", typeof(string));

                // Add sample rows – in a real scenario these would be populated dynamically.
                dataTable.Rows.Add("https://www.example.com");
                dataTable.Rows.Add("https://www.openai.com");
                dataTable.Rows.Add("https://www.github.com");

                // Use WorkbookDesigner to process smart markers.
                var designer = new WorkbookDesigner(workbook);
                designer.SetDataSource("Links", dataTable);
                designer.Process(); // Replaces the placeholder with actual hyperlinks.

                // Save the workbook.
                string outputPath = "HyperlinksOutput.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
