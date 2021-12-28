using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelDataRead
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            string filepath = String.Empty;
            string filetext = String.Empty;
            OpenFileDialog fileDialog = new OpenFileDialog();
            Nullable<bool> dialogOK = fileDialog.ShowDialog();
            if (dialogOK == true)
            {
                filepath = fileDialog.FileName;
                filetext = System.IO.Path.GetExtension(filepath);
                if(filetext.CompareTo(".xls")==0 || filetext.CompareTo(".xlsx")==0)
                {
                    try
                    {
                        System.Data.DataTable table = new System.Data.DataTable();
                        table = ReadExcel(filepath, filetext);
                        
                        var rowcount = table.Rows.Count;
                        foreach (DataRow row in table.Rows)
                        {
                            var nn = row["F4"].ToString();
                        }
                        dtGrid_Excel.DataContext = table.DefaultView;
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            //  if(fileDialog.ShowDialog()== DialogResult.)


        }

        public System.Data.DataTable ReadExcel(string path, string ext)
        {
            string conn = String.Empty;
            System.Data.DataTable table = new System.Data.DataTable();
            if (ext.CompareTo(".xls") == 0)
            {
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//below 2007
            }
            else
            {
                // HRD  HDR
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0 Xml; HDR = YES';";// 2007 or upper version of excel
            }

            using (OleDbConnection con = new OleDbConnection(conn))
            {
                //r3_12_kagoshima_ika_02
                try
                {
                    con.Open();
                    System.Data.DataTable tExcelShetName = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                    string str = "[" + tExcelShetName.Rows[0]["TABLE_NAME"].ToString().Replace("'", "") + "]";
                    MessageBox.Show(str);
                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from " + str, con);
                    adapter.Fill(table);
                    // var tt = table.DataSet.Tables;
                    
                }
                catch (Exception ex1)
                {
                    MessageBox.Show("" + ex1);
                }
            }
            return table;

        }
    }
}
