using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

using System.Drawing.Imaging;

namespace DevicesSpecification
{
    public partial class Form1 : Form
    {

        public struct ShB_elem
        {
            public ShB_elem(string aNumber, string aName, string aColC, string aColF, int aCount)
            {
                Number = aNumber;
                Name = aName;
                ColC = aColC;
                ColF = aColF;
                Count = aCount;
            }
            public string Number { get; }
            public string Name {get; }
            public string ColC { get; }
            public string ColF { get; }
            public int Count { get; }
        };
        public Dictionary<string, List<ShB_elem>> ListShB1;// = new Dictionary<string, ShB1_elem>();
        public Dictionary<string, List<ShB_elem>> ListShB2;// = new Dictionary<string, ShB1_elem>();

        public struct RP_elem
        {
            public RP_elem(string aName, int aCount)
            {
                Name = aName;
                Count = aCount;
            }
            public string Name { get; }
            public int Count { get; }
        };
        public Dictionary<string, Dictionary <string, List<RP_elem>>> USPD;
        public Dictionary<string, Dictionary<string, Dictionary<string, List<RP_elem>>>> TT;

        public Form1()
        {
            InitializeComponent();
            this.Shown += new System.EventHandler(this.Form1_Shown);            
        }
        private async void Form1_Shown(object sender, EventArgs e)
        {
            ListShB1 = new Dictionary<string, List<ShB_elem>>();
            ListShB2 = new Dictionary<string, List<ShB_elem>>();
            USPD = new Dictionary<string, Dictionary<string, List<RP_elem>>>();
            TT   = new Dictionary<string, Dictionary<string, Dictionary<string, List<RP_elem>>>>();
            //await Task.Run(()=>LoadSettings(Directory.GetCurrentDirectory() + "\\Варианты устройства ТТ, Отвл.xlsx"));
            listBox1.Items.AddRange(ListShB1.Keys.ToArray());
            listBox2.Items.AddRange(ListShB2.Keys.ToArray());
            isLoading(false);
        }

        public void  LoadSettings(string file)
        {
            try
            {
                loging(1, "Загрузка настроек...");
                isLoading(true);
                ListShB1.Clear();
                ListShB2.Clear();
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWB;
                Excel.Worksheet xlSht;
                xlWB = xlApp.Workbooks.Open(file);
                
                xlSht = (Excel.Worksheet)xlWB.Worksheets[1];
                Excel.Range last = xlSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                var arrData = (object[,])xlSht.get_Range("A1", last).Value;

                xlSht = (Excel.Worksheet)xlWB.Worksheets[2];
                last = xlSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                var arrData2 = (object[,])xlSht.get_Range("A1", last).Value;

                xlWB.Close(false);
                xlApp.Quit();

                int rowCount = arrData.GetUpperBound(0);
                int colCount = arrData.GetUpperBound(1);
                if (colCount < 7) throw new Exception("Не верные вхрдные данные ШБ.");

                List<ShB_elem> aList = new List<ShB_elem>();
                string variatName = "";
                for (int i = 1; i <= rowCount; i++)
                {                    
                    string aName = getStringFromXML(arrData[i, 2]);
                    if (aName != "")
                    {
                        if (aName.Contains("Вариант"))
                        {
                            if (i > 2)
                                ListShB1.Add(variatName, aList);
                            variatName = aName;
                            aList = new List<ShB_elem>();
                        }
                        else
                            aList.Add(new ShB_elem(
                                getStringFromXML(arrData[i, 1]),
                                aName,
                                getStringFromXML(arrData[i, 3]),
                                getStringFromXML(arrData[i, 6]),
                                getIntFromXML(arrData[i, 7])));
                    }                    
                }
                ListShB1.Add(variatName, aList);

                rowCount = arrData2.GetUpperBound(0);
                colCount = arrData2.GetUpperBound(1);
                if (colCount < 7) throw new Exception("Не верные вхрдные данные ШБ ответвл.");

                aList = new List<ShB_elem>();
                variatName = "";
                for (int i = 1; i <= rowCount; i++)
                {
                    string aName = getStringFromXML(arrData2[i, 2]);
                    if (aName != "")
                    {
                        if (aName.Contains("Вариант"))
                        {
                            if (i > 2)
                                ListShB2.Add(variatName, aList);
                            variatName = aName;
                            aList = new List<ShB_elem>();
                        }
                        else
                            aList.Add(new ShB_elem(
                                getStringFromXML(arrData2[i, 1]),
                                aName,
                                getStringFromXML(arrData2[i, 3]),
                                getStringFromXML(arrData2[i, 6]),
                                getIntFromXML(arrData2[i, 7])));
                    }
                }
                ListShB2.Add(variatName, aList);
                loging(1, "Файл успешно загружен: " + file + ";");
            }
            catch (Exception ex)
            {
                loging(2, "Ошибка загрузки Excel файла: " + file + " ; " + ex.Message);
            }
            isLoading(false);
        }

        private string getStringFromXML(object data)
        {
            string test = "";
            try { test = data.ToString(); }
            catch { }
            return test;
        }
        private int getIntFromXML(object data)
        {
            int test = 0;
            try { test = Convert.ToInt32(data.ToString()); }
            catch { }
            return test;
        }
        public void loging(int level, string text)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<int, string>(loging), new object[] { level, text });
                return;
            }
            var aColor = Color.Black;
            if (level == 1)
                aColor = Color.Green;
            else if (level == 2)
                aColor = Color.Red;
            string curentTime = DateTime.Now.TimeOfDay.ToString("hh\\:mm\\:ss");
            logBox.AppendText(curentTime + ": " + text + Environment.NewLine, aColor);
        }

        public void isLoading (bool value)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<bool>(isLoading), new object[] { value });
                return;
            }
            pictureBox1.Visible = value;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadFilters();
        }

        public void loadFilters()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWB;
            Excel.Worksheet xlSht;
            xlWB = xlApp.Workbooks.Open("D:\\work\\VisualStudio\\DevicesSpecification\\Дахадаевские РЭС Реестр потребителей1.xlsx");
            //Excel.SlicerCache asd = xlWB.SlicerCaches.Item[3];
            //Object[] myObjArray = new Object[1] {"[Кол-во неопрашиваемых ПУ].[Опорная ПС].&[ПО Уркарах Новая]"};
            //asd.VisibleSlicerItemsList = myObjArray;
            xlSht = (Excel.Worksheet)xlWB.Worksheets[9];
            Excel.Range last = xlSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            var arrData = (object[,])xlSht.get_Range("A1", last).Value;

            xlSht = (Excel.Worksheet)xlWB.Worksheets[10];
            last = xlSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            var arrData2 = (object[,])xlSht.get_Range("A1", last).Value;

            xlWB.Close(false);
            xlApp.Quit();

            int rowCount = arrData.GetUpperBound(0);
            int colCount = arrData.GetUpperBound(1);
            //if (colCount < 7) throw new Exception("Не верные вхрдные данные ШБ.");
            
            List<string> ObjName = new List<string>();
            for (int i = 2; i <= colCount; i++)
            {
                string jbvNameElem = getStringFromXML(arrData[3, i]);
                if (jbvNameElem.Contains("Вариант"))
                    ObjName.Add(jbvNameElem);
            }

            List<RP_elem> aList = new List<RP_elem>();
            Dictionary<string, List<RP_elem>> city_USPD = new Dictionary<string, List<RP_elem>>();
            string cityName = "";
            for (int i = 4; i <= rowCount; i++)
            {
                string aName = getStringFromXML(arrData[i, 1]);
                if (aName != "")
                {
                    if (aName.Contains("ПС"))
                    {
                        if (i > 4)
                            USPD.Add(cityName, city_USPD);
                        cityName = aName;
                        city_USPD = new Dictionary<string, List<RP_elem>>();
                    }
                    else
                    {
                        aList = new List<RP_elem>();
                        for (int j = 2; j <= ObjName.Count+1; j++)
                        {
                            int aCount = getIntFromXML(arrData[i, j]);
                            if (aCount > 0)
                                aList.Add(new RP_elem(ObjName[j-2], aCount));
                        }
                        city_USPD.Add(aName, aList);
                    }
                }
            }

            rowCount = arrData2.GetUpperBound(0);
            colCount = arrData2.GetUpperBound(1);
            //if (colCount < 7) throw new Exception("Не верные вхрдные данные ШБ.");

            ObjName = new List<string>();
            for (int i = 2; i <= colCount; i++)
            {
                string jbvNameElem = getStringFromXML(arrData2[3, i]);
                if (jbvNameElem.Contains("/"))
                    ObjName.Add(jbvNameElem);
            }

            aList = new List<RP_elem>();
            Dictionary<string, Dictionary<string, List<RP_elem>>> city_USPD_TT = new Dictionary<string, Dictionary<string, List<RP_elem>>>();
            Dictionary<string, List<RP_elem>> fider_USPD = new Dictionary<string, List<RP_elem>>();
            cityName = "";
            string fideName = "";
            for (int i = 4; i <= rowCount; i++)
            {
                string aName = getStringFromXML(arrData2[i, 1]);
                if (aName.Contains("Итого")) break;
                if (aName != "")
                {
                    if (aName.Contains("ПС"))
                    {
                        if (i > 4)
                        {
                            //fideName = aName;
                            city_USPD_TT.Add(fideName, fider_USPD);
                            fider_USPD = new Dictionary<string, List<RP_elem>>();
                            TT.Add(cityName, city_USPD_TT);
                        }
                            
                        cityName = aName;
                        city_USPD_TT = new Dictionary<string, Dictionary<string, List<RP_elem>>>();
                    }
                    else if (aName.Contains("Фидер") || aName.Contains("того"))
                {
                    if (aName.Contains("того")) aName = "Фидер №1";
                    if (fider_USPD.Count > 0)
                        city_USPD_TT.Add(fideName, fider_USPD);
                    fideName = aName;
                    fider_USPD = new Dictionary<string, List<RP_elem>>();
                }
                else {
                        aList = new List<RP_elem>();
                        for (int j = 2; j <= ObjName.Count + 1; j++)
                        {
                            int aCount = getIntFromXML(arrData2[i, j]);
                            if (aCount > 0)
                                aList.Add(new RP_elem(ObjName[j - 2], aCount));
                        }
                        fider_USPD.Add(aName, aList);
                    }
                }
            }
            city_USPD_TT.Add(fideName, fider_USPD);
            TT.Add(cityName, city_USPD_TT);

        }
    }
    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;

            box.SelectionColor = color;
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
            box.SelectionStart = box.Text.Length;
            box.ScrollToCaret();
        }
    }
}