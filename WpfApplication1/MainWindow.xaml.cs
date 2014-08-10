using System;
using System.Collections.Generic;
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
using Microsoft.Office.Interop;

namespace WpfApplication1
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Window1 win = new Window1();
            win.Show();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            readexcel();
            killAllExcel();
        }

        public bool readexcel()
        {
             String path = AppDomain.CurrentDomain.BaseDirectory;
             String filename = "deadlines.xlsx";
             String file_dir = path + filename;
             Microsoft.Office.Interop.Excel.Application excel_read = new Microsoft.Office.Interop.Excel.Application();
             try
             {
                 Microsoft.Office.Interop.Excel.Workbook exbook = excel_read.Workbooks.Open(file_dir);
                 if (exbook == null)
                 {
                     MessageBox.Show("read error");
                 }
                 Microsoft.Office.Interop.Excel.Worksheet exsheet = exbook.Sheets[1];
                 int row_number = exbook.ActiveSheet.UsedRange.Rows.Count;
                 int i = 0;
                 while (i <= row_number)
                 {
                     String test  = exsheet.Cells.GetType().ToString();
                     //if(exsheet.Cells.GetType().ToString() == double)
                     //{
                     //    DateTime time_scan = DateTime.FromOADate(double.Parse(exsheet.Cells[i, 1]));
                     //}
                     //else if(exsheet.Cells.GetType() == DateTime)
                     //    DateTime timescan = DateTime.Parse(exsheet.Cells[i, 1].ToString())
                     i++;
                     //if (time_scan < DateTime.Now) continue;
                     //else
                     //{
                     //    MessageBox.Show(time_scan.ToString());
                     //    break;
                     //}
                 }
                 TimeSpan cur_day = new TimeSpan(DateTime.Now.Ticks);
                 i = 0 ;
                 MessageBox.Show(exsheet.Cells[2, 1].Text);
                 DateTime event1_dt =
                     Convert.ToDateTime(exsheet.Cells[i + 1 , 1].Text);
                 TimeSpan event1_ts = new TimeSpan(event1_dt.Ticks);
                 event1.Text = exsheet.Cells[i + 1, 2].Text.ToString();
//                 const String event1_text = (const) exsheet.Cells[ i + 1, 2];
                 time1.Text = event1_ts.Subtract(cur_day).Duration().Days.ToString();
                 DateTime event2_dt = Convert.ToDateTime(exsheet.Cells[i + 2, 1].Text);
                 TimeSpan event2_ts = new TimeSpan(event2_dt.Ticks);
                 event2.Text = exsheet.Cells[i + 2, 2].Text.ToString();
                 time2.Text = event2_ts.Subtract(cur_day).Duration().Days.ToString();
                 DateTime event3_dt = Convert.ToDateTime(exsheet.Cells[i + 3, 1].Text);
                 TimeSpan event3_ts = new TimeSpan(event1_dt.Ticks);
                 event3.Text = exsheet.Cells[i + 3, 2].Text.ToString();
                 time3.Text = event3_ts.Subtract(cur_day).Duration().Days.ToString();
                 DateTime event4_dt = Convert.ToDateTime(exsheet.Cells[i + 4, 1].Text);
                 TimeSpan event4_ts = new TimeSpan(event1_dt.Ticks);
                 event4.Text = exsheet.Cells[i + 4, 2].Text.ToString();
                 time4.Text = event4_ts.Subtract(cur_day).Duration().Days.ToString();
                 DateTime event5_dt = Convert.ToDateTime(exsheet.Cells[i + 5, 1].Text);
                 TimeSpan event5_ts = new TimeSpan(event1_dt.Ticks);
                 event5.Text = exsheet.Cells[i + 5, 2].Text.ToString();
                 time5.Text = event5_ts.Subtract(cur_day).Duration().Days.ToString();
                 return true;
             }
            catch(Exception e)
             {
                 MessageBox.Show(e.ToString());
                 return false;
             }
            
        }


        public bool killAllExcel()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    //释放COM组件，其实就是将其引用计数减一

                    foreach (System.Diagnostics.Process process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                    {
                        if (process.CloseMainWindow() == false)
                        {
                            process.Kill();
                        }
                    }
                    excelApp = null;
                } return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
            return true;

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e) 
        {
            if (e.LeftButton == MouseButtonState.Pressed) 
            { 
                DragMove(); 
            } 
        }

    }
}
