using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;

namespace ReadExcel
{
    public class ExcelReader
    {
        public string FileName { get; set; }

        public ExcelReader(string fileName)
        {
            FileName = fileName;
        }

        public void WriteContent()
        {
            var content = GetContent("Orders");

            var q = content.Select(x => new
            {
                // ID	Product Name	Customer Name	Country City	Order Date	Unit Price	Quantity

                //ID = x.Field<string>("ID"),
                //Product = x.Field<string>("Product Name"),
                //Customer = x.Field<string>("Customer Name"),
                //Country = x.Field<string>("Country"),
                //City = x.Field<string>("City"),
                //Date = x.Field<string>("Date"),
                //Unit = x.Field<string>("Unit"),
                //Price = x.Field<string>("Price"),
                //Qty = x.Field<string>("Quantity"),

                ID = x.Field<string>(0),
                Product = x.Field<string>(1),
                Customer = x.Field<string>(2),
                Country = x.Field<string>(3),
                City = x.Field<string>(4),
                Date = x.Field<string>(5),
                UnitPrice = x.Field<string>(6),
                Qty = x.Field<string>(7),
            });

            Console.WriteLine("ID	Product Customer    Date    Quantity");
            q.Skip(1)
             .Where(x => !string.IsNullOrEmpty(x.Product) && x.Product.Length < 10 && !string.IsNullOrEmpty(x.Customer) && x.Customer.Length < 10)
             .OrderBy(x => x.ID.Length)
             .ThenBy(x => x.ID)
             .Take(20)
             .ToList()
             .ForEach(x => Console.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}", x.ID, x.Product, x.Customer, x.Date, x.Qty)));

        }

        private EnumerableRowCollection<DataRow> GetContent(string sheetName)
        {
            // Connection String
            var connstring = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + FileName + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; 

            using (var conn = new OleDbConnection(connstring))
            {
                conn.Open();

                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    // Get All Sheets Name
                    var sheetNames = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {null, null, null, "Table"});

                    // Get the First Sheet Name
                    sheetName = sheetNames.Rows[0][2].ToString();
                }

                // Make sure sheetname ends with '$'
                if (!sheetName.EndsWith("$")) sheetName += "$";

                // Query String 
                var sql = string.Format("SELECT * FROM [{0}]", sheetName);
                var adapter = new OleDbDataAdapter(sql, connstring);
                var dataSet = new DataSet();
                adapter.Fill(dataSet);

                return dataSet.Tables[0].AsEnumerable();
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}