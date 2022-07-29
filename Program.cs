using System.Data;
using ClosedXML.Excel;


class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Processing............................................");
        string filePath = @"C:\Users\Usman Ali\source\repos\ExcelFileReader\bin\Debug\data.xlsx";
        DataTable dt = GetExcelDataTable(filePath);
        List<string> final = new List<string>();
        foreach (DataRow dataRow in dt.Rows)
        {

            //for (int i = 0; i < dataRow.ItemArray.Length; i++)
            //{
                final.Add($"add address ={dataRow.ItemArray[1]} comment = Asset#{dataRow.ItemArray[0]} interface=vlan20-system mac-address={dataRow.ItemArray[2]}");
              
            //  Console.WriteLine(final);
            //}

            //foreach (var item in dataRow.ItemArray)
            //{
            //    Console.WriteLine(item);
            //}
        }
        GenerateDesiredFile(final);
        Console.WriteLine("File has been created! Please check output file path ");
    }

    public static bool GenerateDesiredFile(List<string> data)
    {
        string fileName = @"C:\Users\Usman Ali\source\repos\ExcelFileReader\bin\Debug\sampleData.txt";

        try
        {
            // Check if file already exists. If yes, delete it.     
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            // Create a new file     
            using (StreamWriter sw = File.CreateText(fileName))
            {
                sw.WriteLine("IP Address file: {0}", DateTime.Now.ToString());
                sw.WriteLine("Author: Usman Ali");
                sw.WriteLine("-----------------START----------------------");
                foreach (var item in data)
                {
                    sw.WriteLine(item);
                }
                sw.WriteLine("-----------------END----------------------");
            }
            return true;
        }
        catch (Exception Ex)
        {
            Console.WriteLine(Ex.ToString());
            return false;
        }
    }

    public static DataTable GetExcelDataTable(string filePath)
    {
        DataTable dt = new DataTable();
        using (XLWorkbook workBook = new XLWorkbook(filePath))
        {
            IXLWorksheet workSheet = workBook.Worksheet(1);
            bool firstRow = true;
            foreach (IXLRow row in workSheet.Rows())
            {
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    dt.Rows.Add();
                    int i = 0;
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }
            }
        }

        return dt;
    }
}































//using System;
//using System.Data;
//using System.Data.OleDb;


//namespace CSharpReadExcelFile
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            //this is the connection string which has OLDB 4.0 Connection and Source URL of file
//            //use HDR=YES if first excel row contains headers, HDR=NO means your excel's first row is not headers and it's data.
//            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\\Sample1.xls; Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'";


//            // Create the connection object
//            OleDbConnection oledbConn = new OleDbConnection(connString);
//            try
//            {
//                // Open connection
//                oledbConn.Open();

//                // Create OleDbCommand object and select data from worksheet Sample-spreadsheet-file
//                //here sheet name is Sample-spreadsheet-file, usually it is Sheet1, Sheet2 etc..
//                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sample-spreadsheet-file$]", oledbConn);

//                // Create new OleDbDataAdapter
//                OleDbDataAdapter oleda = new OleDbDataAdapter();

//                oleda.SelectCommand = cmd;

//                // Create a DataSet which will hold the data extracted from the worksheet.
//                DataSet ds = new DataSet();

//                // Fill the DataSet from the data extracted from the worksheet.
//                oleda.Fill(ds, "Employees");

//                //loop through each row
//                foreach (var m in ds.Tables[0].DefaultView)
//                {
//                    Console.WriteLine(((System.Data.DataRowView)m).Row.ItemArray[0] + " " + ((System.Data.DataRowView)m).Row.ItemArray[1] + " " + ((System.Data.DataRowView)m).Row.ItemArray[2]);

//                }

//            }
//            catch (Exception e)
//            {
//                Console.WriteLine("Error :" + e.Message);
//            }
//            finally
//            {
//                // Close connection
//                oledbConn.Close();
//            }
//        }
//    }
//};      