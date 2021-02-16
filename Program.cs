using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data.Common;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.IO;

namespace Zadanie
{

    class Program
    {
        static void GetWell(ref List<string> wells)//Метод где я получаю из бд список скважин.
        {
            string prov = "";
            int i = 0;
            SQLiteConnection conn = new SQLiteConnection("Data Source=C:\\DB\\WellsData4Test.db;");
            conn.Open();
                SQLiteCommand command = new SQLiteCommand("SELECT WELL_NAME FROM 'TBL';", conn);
                SQLiteDataReader well = command.ExecuteReader();
                foreach (DbDataRecord w in well)
                {
                if (i == 0)
                {
                    wells.Add(w["WELL_NAME"].ToString());
                    prov = w["WELL_NAME"].ToString();
                    i++;
                }
                else
                {
                    if(w["WELL_NAME"].ToString() != prov)
                    {
                        wells.Add(w["WELL_NAME"].ToString());
                        prov = w["WELL_NAME"].ToString();
                    }
                }
                }
        }
        static void GetQOilWell(string s ,ref List<string> Q_OIL)//Потом с помощью этого метода я в цикле с помощью списка скважин, получаю список нефти скважины, в дни когда скважина работает
        {
            SQLiteConnection conn = new SQLiteConnection("Data Source=C:\\DB\\WellsData4Test.db;");
            conn.Open();
            SQLiteCommand command = new SQLiteCommand("SELECT Q_OIL_M3 FROM 'TBL' where WELL_NAME = $name and SS='В работе'", conn);
            command.Parameters.AddWithValue("$name", s);
            SQLiteDataReader oil = command.ExecuteReader();
            foreach (DbDataRecord w in oil)
            {
                Q_OIL.Add(w["Q_OIL_M3"].ToString());
            }
        }
        static void GetFluidWell(string s, ref List<string> Q_FLUID)//Потом с помощью этого метода я в цикле с помощью списка скважин, получаю список жидкости скважины, в дни когда скважина работает
        {
            SQLiteConnection conn = new SQLiteConnection("Data Source=C:\\DB\\WellsData4Test.db;");
            conn.Open();
            SQLiteCommand command = new SQLiteCommand("SELECT Q_FLUID_M3 FROM 'TBL' where WELL_NAME = $name and SS='В работе'", conn);
            command.Parameters.AddWithValue("$name", s);
            SQLiteDataReader fluid = command.ExecuteReader();
            foreach (DbDataRecord w in fluid)
            {
                Q_FLUID.Add(w["Q_FLUID_M3"].ToString());
            }
        }
        static void AddChart()//В этом методе из полученных даных я создаю диаграммы
        {
            List<string> well = new List<string>();
            GetWell(ref well);
            double deb = 0;
            int f = 1;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                ExcelLineChart lineChart = worksheet.Drawings.AddChart("lineChart", eChartType.Line) as ExcelLineChart;
                lineChart.Title.Text = "Дебит нефти";
                List<string> Q_OIL = new List<string>();
                foreach (string w in well)
                {
                    GetQOilWell(w, ref Q_OIL);
                    for (int i = 0; i < Q_OIL.Count; i++)
                    {
                        deb = deb + Convert.ToDouble(Q_OIL[i].Replace('.',','));
                        worksheet.Cells[i+1,f].Value =deb;
                        worksheet.Cells[i+1, 5].Value = i+1;
                    }
                    var range1 = worksheet.Cells[1,f,Q_OIL.Count,f];
                    var range2 = worksheet.Cells[1, 5, Q_OIL.Count, 5];
                    var serias = lineChart.Series.Add(range1, range2);
                    serias.Header = w;
                    serias.TrendLines.Add(eTrendLine.Linear);
                    f++;
                    Q_OIL.Clear();
                    deb = 0;
                }
                deb = 0;
                ExcelLineChart lineChart1 = worksheet.Drawings.AddChart("lineChart1", eChartType.Line) as ExcelLineChart;
                lineChart1.Title.Text = "Дебит жидкости";
                List<string> Q_FLUID = new List<string>();
                foreach (string w in well)
                {
                    GetFluidWell(w, ref Q_FLUID);
                    for (int i = 0; i < Q_FLUID.Count; i++)
                    {
                        deb = deb + Convert.ToDouble(Q_FLUID[i].Replace('.', ','));
                        worksheet.Cells[i + 1, f+1].Value = deb;
                        worksheet.Cells[i + 1, 10].Value = i + 1;
                    }
                    var range1 = worksheet.Cells[1, f+1, Q_FLUID.Count, f+1];
                    var range2 = worksheet.Cells[1, 10, Q_FLUID.Count, 10];
                    var serias1 = lineChart1.Series.Add(range1, range2);
                    serias1.Header = w;
                    serias1.TrendLines.Add(eTrendLine.Exponential);//Почитал про линейные тренды, думал взять логарифмическую или экспоненциальную , но так как при логарафмической он показывает возможность отрицательных данных, остановился на экспоненциальной
                    f++;
                    Q_FLUID.Clear();
                    deb = 0;
                }
                FileInfo fi = new FileInfo(@"C:\Users\Мася кун\Desktop\2.xlsx");//не смог понять как сохранять файл в папку проекта, так что сохранял себе на рабочий стол
                excelPackage.SaveAs(fi);
            }

        }
        static void Main(string[] args)
        {
            AddChart();
            Console.ReadKey();
        }
    }
}
