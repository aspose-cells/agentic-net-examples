// Title: Extract CommandText, CommandType, and ConnectionString from worksheet DBConnectionInfo objects with Aspose.Cells for .NET
// Common Searches: c# aspocells read DbConnectionInfo command text from Excel worksheet | how to get command type of DB connections in an Aspose.Cells workbook | using reflection to enumerate DbConnectionInfo collection in Aspose.Cells .NET | retrieve connection strings from each sheet's DB connection in Aspose.Cells | list all DB connection details of worksheets with Aspose.Cells API
// Tags: Aspose.Cells enumerate DbConnectionInfo collection | extract command text from Excel DB connection | retrieve commandtype using Aspose.Cells .NET | read connectionstring from worksheet DB connections | reflection based DB connection inspection Aspose.Cells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet, and uses reflection to access the DbConnectionInfos property. For every DbConnectionInfo found, it extracts the Command (CommandText), CommandType, and ConnectionString, then prints these details to the console while gracefully handling sheets without DB connections or missing properties.
    class Program
    {
        static void Main()
        {
            const string filePath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(filePath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    try
                    {
                        // Use reflection to check for the presence of DbConnectionInfos
                        PropertyInfo dbConnProp = sheet.GetType().GetProperty("DbConnectionInfos");
                        if (dbConnProp != null)
                        {
                            // Retrieve the collection (type is DbConnectionInfoCollection)
                            var dbConnCollection = dbConnProp.GetValue(sheet, null);
                            if (dbConnCollection != null)
                            {
                                // Get the enumerator for the collection
                                var enumerator = ((System.Collections.IEnumerable)dbConnCollection).GetEnumerator();
                                while (enumerator.MoveNext())
                                {
                                    // Each item is a DbConnectionInfo; use dynamic to access its members
                                    dynamic dbConn = enumerator.Current;
                                    string commandText = dbConn.Command;
                                    var commandType = dbConn.CommandType;
                                    string connectionInfo = dbConn.ConnectionString;

                                    Console.WriteLine($"Worksheet: {sheet.Name}");
                                    Console.WriteLine("DBConnection:");
                                    Console.WriteLine($"  CommandText   : {commandText}");
                                    Console.WriteLine($"  CommandType   : {commandType}");
                                    Console.WriteLine($"  ConnectionInfo: {connectionInfo}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Worksheet: {sheet.Name} has no DB connections.");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Worksheet: {sheet.Name} does not support DB connections in this Aspose.Cells version.");
                        }
                    }
                    catch (Exception exSheet)
                    {
                        Console.WriteLine($"Error processing worksheet '{sheet.Name}': {exSheet.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
