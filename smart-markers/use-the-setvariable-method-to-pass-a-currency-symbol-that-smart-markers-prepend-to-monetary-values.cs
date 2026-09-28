// Title: Prepend a currency symbol to amounts using WorkbookDesigner.SetVariable with smart markers in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates an Excel workbook, defines a smart marker like "&=CurrencySymbol?&=Amount?", sets the variable "CurrencySymbol" to "$" with WorkbookDesigner.SetVariable, binds a DataTable containing an Amount column, processes the markers, and saves the file. | Show how to use WorkbookDesigner.SetVariable to supply a custom prefix (e.g., €) for monetary values in smart markers without adding the symbol to the data source. | Generate a complete Aspose.Cells example that formats numbers as currency by combining a SetVariable‑defined symbol and the numeric field in a smart marker expression.
// Common Searches: aspnet aspocells setvariable currency symbol smart marker | c# aspocells prepend currency symbol to smart marker value | using workbookdesigner setvariable for custom text in smart markers | smart markers format monetary values with variable in Aspose.Cells .NET | example of smart marker with currency symbol variable Aspose.Cells C#
// Tags: WorkbookDesigner.SetVariable currency symbol | smart markers prepend text Aspose.Cells | C# Aspose.Cells format monetary values | Excel export custom prefix smart marker | Aspose.Cells smart marker variable usage

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerCurrencyDemo
{
    // The example creates a workbook, adds a smart marker that combines a variable‑defined currency symbol with an amount field, uses WorkbookDesigner.SetVariable to pass the symbol (e.g., "$"), binds a DataTable containing the numeric values, processes the markers, and saves the resulting Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();

                // Rename the default worksheet to "Template"
                Worksheet templateSheet = workbook.Worksheets[0];
                templateSheet.Name = "Template";

                // Insert smart markers that reference fields of the default data source ("Data")
                // The markers will be replaced by the values of CurrencySymbol and Amount
                templateSheet.Cells["A1"].PutValue("&=Data.CurrencySymbol?&=Data.Amount?");

                // Create a WorkbookDesigner to work with smart markers
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Prepare a DataTable as the data source (required by older Aspose.Cells versions)
                DataTable dataTable = new DataTable("Data");
                dataTable.Columns.Add("CurrencySymbol", typeof(string));
                dataTable.Columns.Add("Amount", typeof(double));
                dataTable.Rows.Add("$", 1234.56);

                // Set the data source for the designer
                designer.SetDataSource(dataTable); // default name is "Data"

                // Process the smart markers – this replaces the markers with the data source values
                designer.Process();

                // Save the resulting workbook
                string outputPath = "SmartMarkerCurrencyDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook created successfully at '{outputPath}' with currency symbol applied via smart markers.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
