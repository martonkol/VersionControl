﻿using System;
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

        public Form1()
        {
            InitializeComponent();
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
            //egy string ami tartalmazza a fejleceket
            string[] headers = new string[]
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
                values[counter, 4] = item.Elevator;
                values[counter, 5] = item.NumberOfRooms;
                values[counter, 6] = item.FloorArea;
                values[counter, 7] = item.Price;
                values[counter, 8] = string.Format("={0}/{1}",GetCell(counter+2,7),GetCell(counter+2,8));
                counter++;
            }

            xlSheet.get_Range(
                        GetCell(2, 1),
                        GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;

            FormatTable();
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
        
        void FormatTable()
        {
            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, 9)); //9 helyett headers.length kellene
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
        }

    }
}
