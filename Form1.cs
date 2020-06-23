using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace DevicesSpecification
{
    public partial class Form1 : Form
    {

        public struct ShB_elem
        {
            public ShB_elem(string aNumber, string aName, string aColC, string aColF, int aCount, string aColI)
            {
                Number = aNumber;
                Name = aName;
                ColC = aColC;
                ColF = aColF;
                Count = aCount;
                ColI = aColI;
            }
            public string Number { get; set; }
            public string Name { get; }
            public string ColC { get; set; }
            public string ColF { get; }
            public string ColI { get; }
            public int Count { get; set; }
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
        public Dictionary<string, Dictionary<string, List<RP_elem>>> USPD;
        public Dictionary<string, Dictionary<string, List<RP_elem>>> PU;
        public Dictionary<string, Dictionary<string, Dictionary<string, List<RP_elem>>>> TT;

        public Form1()
        {
            InitializeComponent();
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            this.Shown += new System.EventHandler(this.Form1_Shown);
            //pictureBox1.Dock = DockStyle.Fill;
        }
        private async void Form1_Shown(object sender, EventArgs e)
        {
            ListShB1 = new Dictionary<string, List<ShB_elem>>();
            ListShB2 = new Dictionary<string, List<ShB_elem>>();
            USPD = new Dictionary<string, Dictionary<string, List<RP_elem>>>();
            PU = new Dictionary<string, Dictionary<string, List<RP_elem>>>();
            TT = new Dictionary<string, Dictionary<string, Dictionary<string, List<RP_elem>>>>();
            await Task.Run(() => LoadSettings(Directory.GetCurrentDirectory() + "\\Варианты устройства ТТ, Отвл.xlsx"));
            listBox1.Items.AddRange(ListShB1.Keys.ToArray());
            listBox2.Items.AddRange(ListShB2.Keys.ToArray());
            isLoading(false);
        }

        public void LoadSettings(string file)
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
                                getIntFromXML(arrData[i, 7]),
                                getStringFromXML(arrData[i, 9])));
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
                                getIntFromXML(arrData2[i, 7]),
                                getStringFromXML(arrData[i, 9])));
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

        public void isLoading(bool value)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<bool>(isLoading), new object[] { value });
                return;
            }
            pictureBox1.Visible = value;
            start.Enabled = !value;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            isLoading(true);
            loging(0, "Начало");
            try
            {
                await Task.Run(() => loadFilters());
                await Task.Run(() => GenerateData());
            }
            catch (Exception ex)
            {
                loging(2, ex.Message);
            }
            finally
            {
                isLoading(false);
            }
        }

        public void GenerateData()
        {
            
            loging(0, "Формирование выходных данных");
            List<ShB_elem> result = new List<ShB_elem>();
            foreach (string city in PU.Keys)
            {
                foreach (string fider in PU[city].Keys)
                {
                    result.Add(new ShB_elem("", city + " " + fider, "", "", 0, ""));
                    foreach (RP_elem RP in PU[city][fider])
                    {
                        string varName = RP.Name;
                        int varCount = RP.Count;
                        result.Add(new ShB_elem("", varName, "", "", 0, ""));
                        try
                        {
                            foreach (ShB_elem el in ListShB2[varName.Replace("№", "")])
                            {
                                ShB_elem newEl = el;
                                newEl.Count = newEl.Count * varCount;
                                result.Add(newEl);
                            }
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Не найден вариант " + varName.Replace("№", "") + "в вариантах устройства ТТ." + ex.Message);
                        }
                    }
                    if (TT.ContainsKey(city) & TT[city].ContainsKey(fider))
                        foreach (string varName2 in TT[city][fider].Keys)
                        {
                            int varCount2 = 0;
                            foreach (RP_elem RP in TT[city][fider][varName2])                            
                                varCount2 = varCount2 + RP.Count;
                            
                            foreach (ShB_elem el in ListShB1[varName2])
                            {
                                ShB_elem newEl = el;
                                newEl.Count = newEl.Count * varCount2;
                                result.Add(newEl);
                            }
                            result.RemoveAt(result.Count - 1);
                            int index = 0;
                            foreach (RP_elem RP in TT[city][fider][varName2])
                            {
                                string rpName = RP.Name.Replace("/5", "").Replace("А","");
                                ShB_elem TP = ListShB1[varName2].Last();
                                TP.ColC = "ТОП-0,66 У3 "+ rpName +"/ 5 0,5S";
                                TP.Number = (Int32.Parse(TP.Number) + index).ToString();
                                result.Add(TP);
                            }
                        }
                }
            }

            loging(0, "Формирование выходных данных успешно завершено");

            object[,] arr = new object[result.Count, 9];

            int i = -1;
            foreach (ShB_elem el in result)
            {
                i++;
                arr[i, 0] = el.Number;
                arr[i, 1] = el.Name;
                arr[i, 2] = el.ColC;
                //arr[i, 3] = el.Number;
                //arr[i, 4] = el.Number;
                arr[i, 5] = el.ColF;
                arr[i, 6] = el.Count;
                //arr[i, 7] = el.Number;
                arr[i, 8] = el.ColI;
            }
            
        }

        public void loadFilters()
        {
            TT.Clear();
            USPD.Clear();
            PU.Clear();
            loging(0, "Чтение файла");
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWB;
            Excel.Worksheet xlSht;
            Excel.Range last;
            string fileName = textBox1.Text;
            if (!File.Exists(fileName)) { loging(2, "Проверьте пусть к файлу."); return; }
            xlWB = xlApp.Workbooks.Open(fileName);
            //Excel.SlicerCache asd = xlWB.SlicerCaches.Item[3];
            //Object[] myObjArray = new Object[1] {"[Кол-во неопрашиваемых ПУ].[Опорная ПС].&[ПО Уркарах Новая]"};
            //asd.VisibleSlicerItemsList = myObjArray;

            xlSht = (Excel.Worksheet)xlWB.Worksheets[9];
            last = xlSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            var arrData = (object[,])xlSht.get_Range("A1", last).Value;

            xlSht = (Excel.Worksheet)xlWB.Worksheets[10];
            last = xlSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            var arrData2 = (object[,])xlSht.get_Range("A1", last).Value;

            xlSht = (Excel.Worksheet)xlWB.Worksheets[6];
            last = xlSht.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            var arrData3 = (object[,])xlSht.get_Range("A1", last).Value;

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
                        for (int j = 2; j <= ObjName.Count + 1; j++)
                        {
                            int aCount = getIntFromXML(arrData[i, j]);
                            if (aCount > 0)
                                aList.Add(new RP_elem(ObjName[j - 2], aCount));
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
                    else
                    {
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

            rowCount = arrData3.GetUpperBound(0);
            colCount = arrData3.GetUpperBound(1);

            ObjName = new List<string>();
            for (int i = 2; i <= colCount; i++)
            {
                string jbvNameElem = getStringFromXML(arrData3[3, i]);
                if (jbvNameElem.Contains("№"))
                    ObjName.Add(jbvNameElem);
            }
            cityName = "";
            Dictionary<string, List<RP_elem>> fiderList = new Dictionary<string, List<RP_elem>>();
            for (int i = 4; i <= rowCount; i++)
            {
                string aName = getStringFromXML(arrData3[i, 1]);
                if (aName.Contains("№"))
                {
                    for (int j = 2; j <= ObjName.Count + 1; j++)
                    {
                        int aCount = getIntFromXML(arrData3[i, j]);
                        if (aCount > 0)
                        {
                            if (!fiderList.ContainsKey(ObjName[j - 2])) fiderList.Add(ObjName[j - 2], new List<RP_elem>());
                            fiderList[ObjName[j - 2]].Add(new RP_elem(aName, aCount));
                        }
                    }
                }
                else if (aName.Contains("ПС"))
                {
                    if (i > 4)
                        PU.Add(cityName, fiderList);
                    cityName = aName;
                    fiderList = new Dictionary<string, List<RP_elem>>();
                }
            }
            PU.Add(cityName, fiderList);

            loging(0, "Чтение файла завершено успешно");
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                if (file.Contains(".xls") | file.Contains(".xlsx"))
                {
                    textBox1.Text = file;
                    loging(0, "Добавлен файл " + file);
                }
                else
                {
                    loging(2, "Не корректный тип файла " + file);
                }
            }
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