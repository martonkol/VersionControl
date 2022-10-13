using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; //Excel == : alias
using System.Reflection; //technikai könyvtár
using System.Data.Entity.Migrations.Model;

namespace ExcelExport
{
    public partial class Form1 : Form
    {   
        //lista a Flat elemekből, neve: flats
        List<Flat> flats;
        //peldanyositom az ORM objektumot
        RealEstateEntities context = new RealEstateEntities();

        //Excel változók --
        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;
        //--

        //egy string ami tartalmazza a fejleceket
        public string[] headers;


        public Form1()
        {
            //InitializeComponent();
            LoadData();
            CreateExcel();
        }

        void LoadData()
        {
            //A context elembol a Flat elemeket hozzaadja a Flat elemeket tartalmazo flats listahoz
            flats = context.Flats.ToList();
        }

        void CreateExcel()
        {
            //try-catch blokk
            try
            {
                //excel elinditasa és az application object betoltese
                xlApp = new Excel.Application();

                //uj munkafuzet letrehozasa
                xlWB = xlApp.Workbooks.Add(Missing.Value);

                //uj munkalap letrehozasa
                xlSheet = xlWB.ActiveSheet;

                //tabla letrehozasa fgv
                CreateTable();

                //control atadasa a felhasznalonak
                xlApp.Visible = true;
                xlApp.UserControl = true;
                //xlApp.Save(autoSave());

            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg);

                // Hiba esetén az Excel applikáció bezárása automatikusan
                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }

        }

        //az excel mindig range-be ír (tehát nem soronként)

        void CreateTable()
        {
            headers = new string[]
            {
                 "Kód",
                 "Eladó",
                 "Oldal",
                 "Kerület",
                 "Lift",
                 "Szobák száma",
                 "Alapterület (m2)",
                 "Ár (mFt)",
                 "Négyzetméter ár (Ft/m2)"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, 1+i] = headers[i]; //cells[sor, oszlop]
            }

            //object tipusu 2d tomb az adatok tarolasara
            //a flats lista elemek szamabol (sorok)
            //a headers string tomb elemek szamabol (oszlopok)
            object[,] values = new object[flats.Count, headers.Length];

            int counter = 0;
            foreach (Flat item in flats)
            {
                values[counter, 0] = item.Code;
                values[counter, 1] = item.Vendor;
                values[counter, 2] = item.Side;
                values[counter, 3] = item.District;

                //
                if (item.Elevator)
                {
                    values[counter, 4] = "Van";
                }
                else
                {
                    values[counter, 4] = "Nincs";
                }

                //
                values[counter, 5] = item.NumberOfRooms;
                values[counter, 6] = item.FloorArea;
                values[counter, 7] = item.Price;
                values[counter, 8] = string.Format("={0}/{1}", GetCell(counter + 2, 8), GetCell(counter + 2, 7));
                counter++;
            }

            xlSheet.get_Range(
                        GetCell(2, 1),
                        GetCell(1 + values.GetLength(0), values.GetLength(1))
                        ).Value2 = values;
            //xlSheet.getRange(A2 I33).Value2 = values

            FormatHeader();
            FormatContent();
        }


        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }
        
        void FormatHeader()
        {
            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length)); //9 helyett headers.length kellene
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
        }

        void FormatContent()
        {
            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range contentRange = xlSheet.get_Range(GetCell(2, 1), GetCell(lastRowID, 9));
            contentRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range lastcolumnRange = xlSheet.get_Range(GetCell(2, 9), GetCell(lastRowID, 9));
            lastcolumnRange.Interior.Color = Color.LightGreen;
            //lastcolumnRange.

            Excel.Range firstcolumnRange = xlSheet.get_Range(GetCell(2, 1), GetCell(lastRowID, 1));
            firstcolumnRange.Interior.Color = Color.LightYellow;
            firstcolumnRange.Font.Bold = true;
        }
    }
}
