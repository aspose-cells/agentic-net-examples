// Title: C# – Filter Null or Blank Rows After Smart Marker Processing with Aspose.Cells AutoFilter
// Description: Demonstrates how to bind a list with nullable fields to Aspose.Cells smart markers, process them with WorkbookDesigner, and then use AutoFilter.MatchNonBlanks to hide rows where Name or Score is null or empty before saving the workbook.
// Keywords: Aspose.Cells | C# | Smart markers | AutoFilter | filter null values | exclude blank rows | WorkbookDesigner | UpdateEmptyStringAsNull | Excel report generation | nullable data source
// Common Searches: Aspose.Cells filter rows with null values after smart markers | C# AutoFilter on smart marker output | hide blank rows in Excel using Aspose.Cells | exclude records with empty cells in smart marker report | UpdateEmptyStringAsNull Aspose.Cells example
// Developer Intent: Remove rows that contain null or empty Name or Score fields after expanding smart markers, so the final Excel file contains only complete records.
// Use Cases: Create a clean sales report that lists only customers with both name and score. | Export a dataset to Excel while automatically discarding incomplete entries before analysis. | Prepare chart‑ready data by filtering out rows with missing values after smart marker expansion.
// AI Prompts: Generate C# code using Aspose.Cells WorkbookDesigner and AutoFilter to exclude rows with null Name or Score after processing smart markers. | Explain how UpdateEmptyStringAsNull and AutoFilter.MatchNonBlanks work together to hide blank cells in an Aspose.Cells workbook. | Provide a step‑by‑step tutorial for applying a smart marker filter that removes rows with null values before saving the Excel file.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerFiltering
{
    // Demonstrates how to bind a list with nullable fields to Aspose.Cells smart markers, process them with WorkbookDesigner, and then use AutoFilter.MatchNonBlanks to hide rows where Name or Score is null or empty before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add headers for the data
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Score");

                // Insert smart markers that will be replaced by the data source
                sheet.Cells["A2"].PutValue("&=$Data.Name");
                sheet.Cells["B2"].PutValue("&=$Data.Score");

                // Prepare a strongly‑typed data source (all Score values are nullable)
                var data = new List<DataItem>
                {
                    new DataItem { Name = "Alice",   Score = 85 },
                    new DataItem { Name = "Bob",     Score = null },   // Null score
                    new DataItem { Name = null,      Score = 90 },    // Null name
                    new DataItem { Name = "Charlie",Score = 78 }
                };

                // Set up the WorkbookDesigner
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    // Treat empty strings as null (helps when nulls are represented as empty strings)
                    UpdateEmptyStringAsNull = true
                };

                // Bind the data source to the smart marker group named "Data"
                designer.SetDataSource("Data", data);

                // Process the smart markers – this populates the worksheet with the data
                designer.Process();

                // Apply an AutoFilter to hide rows where either Name or Score is blank/null
                // The range includes the header row and all data rows (4 rows in this example)
                sheet.AutoFilter.Range = "A1:B5";

                // Show only rows with non‑blank values in column A (Name)
                sheet.AutoFilter.MatchNonBlanks(0);
                // Show only rows with non‑blank values in column B (Score)
                sheet.AutoFilter.MatchNonBlanks(1);

                // Refresh the filter to apply the criteria
                sheet.AutoFilter.Refresh();

                // Save the resulting workbook
                string outputPath = "FilteredSmartMarkers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Simple POCO to hold the data; Score is nullable to allow null values.
        private class DataItem
        {
            public string Name { get; set; }
            public int? Score { get; set; }
        }
    }
}
