using System;
using System.Threading;
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
using System.Windows.Shapes;
using Microsoft.Office.Interop;

namespace WpfApplication1
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        DateTime sel_date;
        String eventname;
        String eventdescription;
        Boolean important;
        public Window1()
        {
            InitializeComponent();
        }

        protected void MonthlyCalendar_SelectedChanged(object sender, SelectionChangedEventArgs e)
        {
            DatePicker datepicker1 = new DatePicker();
            sel_date = (DateTime) date_picker.SelectedDate;
            MessageBox.Show(sel_date.ToString());
        }

        protected void RichTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            eventname = Console.ReadLine();

        }


        public void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            eventdescription = Console.ReadLine();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            eventname = Event_Name.Text.ToString();
            eventdescription = Event_Description.Text.ToString();
            SaveDataToExcel(sel_date, eventname, eventdescription, important);
            killAllExcel();
            this.Close();
        }

        public static bool SaveDataToExcel(DateTime date, string event_name, string event_description, Boolean important_event)
        {
            Microsoft.Office.Interop.Excel.Application excel_write = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                String path = AppDomain.CurrentDomain.BaseDirectory;
                String filename = "deadlines.xlsx";
                String file_dir = path + filename;
                excel_write.DisplayAlerts = false;
                Microsoft.Office.Interop.Excel.Workbook exbook = excel_write.Workbooks.Open(file_dir);
                if (exbook == null)
                {
                    MessageBox.Show("no find, create one");
                    excel_write.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet exsheet_tmp = exbook.Sheets[1];
                    exsheet_tmp.SaveAs(file_dir);

                }
                Microsoft.Office.Interop.Excel.Worksheet exsheet = exbook.Sheets[1];
                int row_num = exbook.ActiveSheet.UsedRange.Rows.Count;
                //int row_num = exsheet.Rows.Count;
                MessageBox.Show(row_num.ToString());
                
                exsheet.Cells[row_num + 1, 1] = date.ToShortDateString().ToString();
                exsheet.Cells[row_num + 1, 2] = event_name;
                exsheet.Cells[row_num + 1, 3] = event_description;

                if (important_event == true)
                {
                    for(int i = 1 ; i<4; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range im_rang = (Microsoft.Office.Interop.Excel.Range)exsheet.Cells[row_num + 1, i];
                        im_rang.Interior.ColorIndex = 3;
                       
                    }
                    MessageBox.Show("paint to red");
                }

////                if (important_event == true)
//                {
//                    Microsoft.Office.Interop.Excel.Range range = excel_write.get_Range("A1:A3, E1:G2");
//                    range.Interior.ColorIndex = color;
//                }

                exsheet.SaveAs(file_dir);
                MessageBox.Show("writing finished");

                exbook.Close();


                //excel_write.Application.Workbooks.Add(true);
                //excel_write.Visible = true;
                //if (excel_write == null) { MessageBox.Show("can't open excel"); return false; }
                //excel_write.Workbooks.Add(true);
                //int row = excel_write.Rows.Count;
                //return true;
                
                return true;

            }
            catch (Exception e) {
                MessageBox.Show(e.ToString());
                return true;
            }
            
        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {   
            if((bool) checkBox.IsChecked)
            {
                important = true;
                MessageBox.Show("this is important");
            }
        }

        public bool killAllExcel()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            try{
                if(excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                    //释放COM组件，其实就是将其引用计数减一
                    
                    foreach(System.Diagnostics.Process process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
                    {
                        if(process.CloseMainWindow() == false)
                        {
                            process.Kill();
                        }
                    }
                    excelApp = null;
                } return true;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
            return true;
        
        }
        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
