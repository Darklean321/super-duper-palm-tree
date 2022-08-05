using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.IO;

namespace _ExcelRebuildWF
{
    internal class ExcelActions
    {
        public bool check1 = false;
        public bool check2 = false;


        public static void ReadExcel()
        {
            EXL EX_DATA = new EXL();
            object[,] dataArr = null;

            EXL EX_WRITE = new EXL();
            //object[,] dataArr1 = null;

            var listObjects = new List<ExcelObject>();

            try
            {
                /*OpenFileDialog inputFile = new OpenFileDialog();

                inputFile.Filter = "Файлы Excel (*.xls;*.xlsx)|*.xls;*.xlsx|Все файлы (*.*)|*.*";
                inputFile.FilterIndex = 1;
                inputFile.RestoreDirectory = true;*/

                FolderBrowserDialog inputFolder = new FolderBrowserDialog();

                /*if (inputFile.ShowDialog() != DialogResult.OK)
                    return;*/
                if (inputFolder.ShowDialog() != DialogResult.OK)
                    return;

                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();



                string[] files = new DirectoryInfo($"{inputFolder.SelectedPath}\\").GetFiles("*xls?", SearchOption.AllDirectories)
                    .Select(f => f.FullName).ToArray();
                foreach (string xlFileName in files)
                {
                    //string xlFileName = inputFile.FileName;
                    if (!xlFileName.Contains("_1С"))
                    {
                        Excel.Range Rng;
                        EX_DATA.App = new Excel.Application();
                        EX_DATA.WBs = EX_DATA.App.Workbooks;
                        EX_DATA.WB = EX_DATA.WBs.Open(xlFileName);
                        EX_DATA.Shts = EX_DATA.WB.Worksheets;
                        EX_DATA.Sht = EX_DATA.Shts.Item[1];
                        EX_DATA.cell = EX_DATA.Sht.Cells[1, 1];

                        int iLastRow = EX_DATA.cell[EX_DATA.Sht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
                        int iLastCol = EX_DATA.cell[1, EX_DATA.Sht.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

                        Rng = (Excel.Range)EX_DATA.Sht.Range["A1", EX_DATA.Sht.Cells[iLastRow, iLastCol]];

                        dataArr = (object[,])Rng.Value2;

                        string[] arrCol = new string[iLastCol];

                        List<string> materials = new List<string>();

                        string dec_sep = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;



                        for (int i = 0; i < iLastRow; i++)
                        {
                            listObjects.Add(new ExcelObject());

                            listObjects[i].Header = EX_DATA.cell[i + 1, "A"].Value2;

                            listObjects[i].Наименование = EX_DATA.cell[i + 1, "B"].Value2;

                            listObjects[i].Обозначение = EX_DATA.cell[i + 1, "C"].Value2;

                            if (EX_DATA.cell[i + 1, "D"].Value.ToString() != null)
                            {
                                double number;
                                bool isNumber = double.TryParse(EX_DATA.cell[i + 1, "D"].Value2.ToString().Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number);
                                if (isNumber)
                                {
                                    EX_DATA.cell[i + 1, "D"].NumberFormat = $"#{dec_sep}##0";
                                    EX_DATA.cell[i + 1, "D"].Value2 = number;
                                    listObjects[i].Количество = EX_DATA.cell[i + 1, "D"].Value2;
                                }
                            }

                            listObjects[i].Материал = EX_DATA.cell[i + 1, "E"].Value2;

                            if (listObjects[i].Header == "Сборочные единицы")
                                if (!materials.Contains($"{listObjects[i].Наименование} {listObjects[i].Обозначение}"))
                                {
                                    materials.Add($"{listObjects[i].Наименование} {listObjects[i].Обозначение}");
                                }

                            if (listObjects[i].Header == "Детали")
                            {
                                if (listObjects[i].Материал != null)
                                {
                                    if (listObjects[i].Материал.Contains("Полиамид") && !materials.Contains($"{listObjects[i].Наименование} {listObjects[i].Обозначение}"))
                                    {
                                        materials.Add($"{listObjects[i].Наименование} {listObjects[i].Обозначение}");
                                    }
                                    if ((listObjects[i].Материал == "" || listObjects[i].Материал == null) && !materials.Contains($"{listObjects[i].Наименование} {listObjects[i].Обозначение}"))
                                    {
                                        materials.Add($"{listObjects[i].Наименование} {listObjects[i].Обозначение}");
                                    }
                                    if (!listObjects[i].Материал.Contains("Полиамид") && !listObjects[i].Материал.Contains("Белый пластик")
                                        && !materials.Contains(listObjects[i].Материал))
                                    {
                                        materials.Add(listObjects[i].Материал);
                                    }
                                }
                                else if (listObjects[i].Материал == null && listObjects[i].Наименование != null &&
                                    !materials.Contains($"{listObjects[i].Наименование} {listObjects[i].Обозначение}"))
                                {
                                    materials.Add($"{listObjects[i].Наименование} {listObjects[i].Обозначение}");
                                }
                            }

                            if (listObjects[i].Header == "Материалы")
                            {
                                if (!listObjects[i].Наименование.Contains("Наполнение") && !materials.Contains(listObjects[i].Наименование))
                                {
                                    materials.Add(listObjects[i].Наименование);
                                }
                                if (listObjects[i].Наименование.Contains("Наполнение") && !materials.Contains($"{listObjects[i].Наименование} {listObjects[i].Обозначение}"))
                                {
                                    materials.Add($"{listObjects[i].Наименование} {listObjects[i].Обозначение}");
                                }
                            }

                            if (listObjects[i].Header == "Стандартные изделия" || listObjects[i].Header == "Прочие изделия")
                                if (!materials.Contains(listObjects[i].Наименование))
                                {
                                    materials.Add(listObjects[i].Наименование);
                                }

                            if (EX_DATA.cell[i + 1, "K"].Value2 != null)
                            {
                                double number;
                                bool isNumber = double.TryParse(EX_DATA.cell[i + 1, "K"].Value2.ToString().Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number);
                                if (isNumber)
                                {
                                    EX_DATA.cell[i + 1, "K"].NumberFormat = $"#{dec_sep}##0";
                                    EX_DATA.cell[i + 1, "K"].Value2 = number;
                                    listObjects[i].Размер = EX_DATA.cell[i + 1, "K"].Value2;
                                }
                            }

                            if (EX_DATA.cell[i + 1, "AA"].Value2 != null && i != 0)
                            {
                                double number;
                                bool isNumber = double.TryParse(EX_DATA.cell[i + 1, "AA"].Value2.ToString().Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out number);
                                if (isNumber)
                                {
                                    EX_DATA.cell[i + 1, "AA"].NumberFormat = $"#{dec_sep}##0";
                                    EX_DATA.cell[i + 1, "AA"].Value2 = number;
                                    listObjects[i].Вес = EX_DATA.cell[i + 1, "AA"].Value2;
                                }
                            }


                        }
                        EX_DATA.App.DisplayAlerts = false;
                        EX_DATA.WBs.Close();
                        EX_DATA.App.Quit();
                        EX_DATA.App.DisplayAlerts = true;

                        Marshal.ReleaseComObject(EX_DATA.cell);
                        Marshal.ReleaseComObject(EX_DATA.Sht);
                        Marshal.ReleaseComObject(EX_DATA.Shts);
                        Marshal.ReleaseComObject(EX_DATA.WB);
                        Marshal.ReleaseComObject(EX_DATA.WBs);
                        Marshal.ReleaseComObject(EX_DATA.App);

                        EX_WRITE.App = new Excel.Application();
                        EX_WRITE.App.SheetsInNewWorkbook = 1;
                        EX_WRITE.WB = EX_WRITE.App.Workbooks.Add();
                        EX_WRITE.Sht = EX_WRITE.WB.Worksheets[1];



                        for (int i = 0; i < materials.Count; i++)
                        {
                            if (materials[i].Contains("МАТЕРИАЛ"))
                            {
                                materials.RemoveAt(i);
                            }
                        }
                        for (int i = 0; i < materials.Count; i++)
                        {
                            if (materials[i] == "")
                            {
                                materials.RemoveAt(i);
                            }
                        }

                        int excelIndex = 0;

                        materials.Sort();

                        foreach (string material in materials)
                        {
                            double materialSumm = 0;
                            int count = 0;
                            foreach (var listObject in listObjects)
                            {
                                string materialHeader = listObject.Header;
                                string nameComparer = $"{listObject.Наименование} {listObject.Обозначение}";

                                if (listObject.Header == "Сборочные единицы")
                                {
                                    if (nameComparer == material)
                                    {
                                        materialSumm += listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(9)}{excelIndex + 1}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(9)}{excelIndex + 1}"].Value2 = count;
                                    }
                                }

                                if (listObject.Header == "Детали")
                                {
                                    if (listObject.Материал == material && (listObject.Материал.Contains("Прокат") ||
                                        listObject.Материал.Contains("Лист") || listObject.Материал.Contains("Рулон")))
                                    {
                                        materialSumm += listObject.Вес * listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(2)}{excelIndex + 1}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(2)}{excelIndex + 1}"].Value2 = count;
                                    }
                                    if (listObject.Материал == material && (listObject.Материал.Contains("Ригель") ||
                                        listObject.Материал.Contains("Стойка")))
                                    {
                                        materialSumm += listObject.Размер / 1000 * listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{excelIndex + 1}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{excelIndex + 1}"].Value2 = count;
                                    }
                                    if (nameComparer == material && (listObject.Материал == null || listObject.Материал.Contains("Полиамид") ||
                                        listObject.Материал == ""))
                                    {
                                        materialSumm += listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex + 1}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex + 1}"].Value2 = count;
                                    }
                                    if (listObject.Материал == material && (listObject.Материал == "" || listObject.Материал == null))
                                    {
                                        materialSumm += listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(5)}{excelIndex + 1}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(5)}{excelIndex + 1}"].Value2 = count;
                                    }
                                    if (listObject.Материал == material && listObject.Наименование.Contains("Шина"))
                                    {
                                        materialSumm += listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(6)}{excelIndex + 1}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(6)}{excelIndex + 1}"].Value2 = count;
                                    }
                                }

                                else if (listObject.Header == "Материалы")
                                {
                                    if (listObject.Наименование == material)
                                    {
                                        materialSumm += listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(7)}{excelIndex + 1}"].Interior.Color =
                                                System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(7)}{excelIndex + 1}"].Value2 = count;
                                    }
                                    else if ($"{listObject.Обозначение} {listObject.Наименование}" == material)
                                    {
                                        materialSumm += listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(8)}{excelIndex + 1}"].Interior.Color =
                                                System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Pink);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(8)}{excelIndex + 1}"].Value2 = count;
                                    }
                                }

                                else if (listObject.Header == "Стандартные изделия" || listObject.Header == "Прочие изделия")
                                {
                                    if (listObject.Наименование == material)
                                    {
                                        materialSumm += listObject.Количество;
                                        EX_WRITE.Sht.Range[$"{GetLetter(9)}{excelIndex + 1}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
                                        count++;
                                        EX_WRITE.Sht.Range[$"{GetLetter(9)}{excelIndex + 1}"].Value2 = count;
                                    }
                                }
                            }

                            EX_WRITE.Sht.Range[$"{GetLetter(0)}{excelIndex + 1}"].Value2 = material;


                            EX_WRITE.Sht.Range[$"{GetLetter(1)}{excelIndex + 1}"].NumberFormat = "#,##0.00";
                            if (EX_WRITE.Sht.Range[$"{GetLetter(1)}{excelIndex + 1}"].Value2 != null)
                                EX_WRITE.Sht.Range[$"{GetLetter(1)}{excelIndex + 1}"].Value2 = (double)EX_WRITE.Sht.Range[$"{GetLetter(1)}{excelIndex + 1}"].Value2;

                            EX_WRITE.Sht.Range[$"{GetLetter(1)}{excelIndex + 1}"].Value2 = materialSumm.ToString();//.Replace(" ", ".").Replace(",", ".");
                            excelIndex++;
                        }

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 = "(Рулоны и листы металла) Вес * Количество";

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 = "(Ригели, стойки) Размер / 1000 * Количество";

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 =
                            "(По наименованию и обозначению: Полиамид или материал пуст или null) Количество";

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 = "(По материалу: Материал пуст или null) Количество";

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Cyan);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 = "(Шина) Количество";

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 = "(Наименование = Материал) Количество";

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Pink);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 = "(Наименование + Обозначение = Материал) Количество";

                        EX_WRITE.Sht.Range[$"{GetLetter(3)}{++excelIndex}"].Interior.Color =
                                            System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Violet);
                        EX_WRITE.Sht.Range[$"{GetLetter(4)}{excelIndex}"].Value2 = "(Сборочные единицы, Метизы и Покупные) Количество";


                        if (xlFileName.Contains(".xls") && !xlFileName.Contains(".xlsx"))
                        {
                            EX_WRITE.WB.SaveAs(xlFileName.Substring(0, xlFileName.Length - 4) + "_1С.xlsx");
                        }
                        else if (xlFileName.Contains(".xlsx"))
                        {
                            EX_WRITE.WB.SaveAs(xlFileName.Substring(0, xlFileName.Length - 5) + "_1С.xlsx");
                        }

                        EX_WRITE.App.Quit();

                        Marshal.ReleaseComObject(EX_WRITE.Sht);
                        Marshal.ReleaseComObject(EX_WRITE.WB);
                        Marshal.ReleaseComObject(EX_WRITE.App);
                    }
                }
                stopWatch.Stop();
                // Get the elapsed time as a TimeSpan value.
                TimeSpan ts = stopWatch.Elapsed;

                // Format and display the TimeSpan value.
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                    ts.Hours, ts.Minutes, ts.Seconds,
                    ts.Milliseconds / 10);

                MessageBox.Show("Выполнение программы завершено\nRunTime: " + elapsedTime, "Разбор по фрагментам(сборка)");
            }
            finally
            {
                //освобождаем память, занятую объектами
                Marshal.ReleaseComObject(EX_DATA.cell);
                Marshal.ReleaseComObject(EX_DATA.Sht);
                Marshal.ReleaseComObject(EX_DATA.Shts);
                Marshal.ReleaseComObject(EX_DATA.WB);
                Marshal.ReleaseComObject(EX_DATA.WBs);
                Marshal.ReleaseComObject(EX_DATA.App);

                Marshal.ReleaseComObject(EX_WRITE.Sht);
                Marshal.ReleaseComObject(EX_WRITE.WB);
                Marshal.ReleaseComObject(EX_WRITE.App);
            }
        }

        public struct EXL
        {
            public Excel.Application App;
            public Excel.Workbooks WBs;
            public Excel.Workbook WB;
            public Excel.Sheets Shts;
            public Excel.Worksheet Sht;
            public Excel.Range cell;
        }

        static private string GetLetter(int nn)
        {
            string p1;

            int n2 = nn / 26;
            if (n2 > 0)
            {
                p1 = ((char)((int)('A') + n2 - 1)).ToString() + ((char)((int)('A') + nn - n2 * 26)).ToString();
            }
            else
            {
                p1 = ((char)((int)('A') + nn)).ToString();
            }

            return p1;
        }
    }
}
