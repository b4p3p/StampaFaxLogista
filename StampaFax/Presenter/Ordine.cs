
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace StampaFax.Presenter
{
    internal class Ordine
    {
        private const int _RIGA_CODICE = 4;

        List<Record> lstRecord;
        RecTotale[] totaliFogli = new RecTotale[2];

        public Ordine()
        {
            lstRecord = new List<Record>();

            string path = new FileInfo(Assembly.GetExecutingAssembly().Location).DirectoryName;

            FileInfo fi = new FileInfo(path + "\\TABACCHI.xlsm");

            if (fi.Exists == false) 
                throw new FileNotFoundException("File TABACCHI.xlsm non trovato");

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fi.FullName);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["temp_ordine"];

            for (int i = _RIGA_CODICE + 1; i < _RIGA_CODICE + 24 * 4; i++)
            {
                string codice = xlWorksheet.Cells[i, 1].Text;
                string grammiCad = xlWorksheet.Cells[i, 2].Text;
                string nome = xlWorksheet.Cells[i, 3].Text;
                string stecche = xlWorksheet.Cells[i, 4].Text;
                string chili = xlWorksheet.Cells[i, 5].Text;
                string grammi = xlWorksheet.Cells[i, 6].Text;

                if (codice == "" && grammiCad == "") break;

                Record r = new Record()
                {
                    codice = codice,
                    grammiCad = grammiCad,
                    stecche = stecche,
                    chili = chili,
                    grammi = grammi
                };

                lstRecord.Add(r);
            }

            //copio i totali dal foglio
            for (int i = 0; i < 2; i++)
            {
                string chili = xlWorksheet.Cells[1 + i, 5].Text;
                string grammi = xlWorksheet.Cells[1 + i, 6].Text;

                chili = chili.PadLeft(6, ' ');
                chili = chili.Insert(3, " ");
                string chiliConv = "";
                foreach (Char c in chili)
                {
                    chiliConv += c.ToString();
                    chiliConv += " ";
                }

                if (grammi == "0") grammi = "000";

                grammi = grammi.PadLeft(3, ' ');
                string grammiConv = "";
                foreach (Char c in grammi)
                {
                    grammiConv += c.ToString();
                    grammiConv += " ";
                }

                RecTotale rt = new RecTotale()
                {
                    chili = chiliConv,
                    grammi = grammiConv
                };

                totaliFogli[i] = rt;
            }

            xlWorkbook.Close();

        }

        internal int getNumeroRecord()
        {
            return lstRecord.Count;
        }

        internal Record getOrdine(int i)
        {
            return lstRecord[i];
        }

        internal string getChiliTot(int foglio)
        {
            return totaliFogli[foglio - 1].chili;
        }

        internal string getGrammiTot(int foglio)
        {
            return totaliFogli[foglio - 1].grammi;
        }


        private class RecTotale
        {
            public string chili { get; set; }
            public string grammi { get; set; }
        }
    }
}
