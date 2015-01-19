using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StampaFax.Presenter
{
    internal class Record
    {
        public string codice { get; set; }
        internal string StampaCodice()
        {
            string ris = "";

            //if ( codice.Length != 4 ) codice = codice.PadLeft( 4 , '0' );

            foreach (Char c in codice)
            {
                ris += c.ToString();
                ris += " ";
            }

            return ris;

        }
        public string chili { get; set; }
        internal string StampaChili()
        {
            string ris = "";

            if (chili.Length != 3) chili = chili.PadLeft(3, ' ');

            foreach (Char c in chili)
            {
                ris += c.ToString();
                ris += " ";
            }

            return ris;
        }
        public string grammi { get; set; }
        internal string StampaGrammi()
        {
            string ris = "";

            //if (grammi.Length != 3) grammi = grammi.PadLeft(3, '0');

            foreach (Char c in grammi)
            {
                ris += c.ToString();
                ris += " ";
            }

            return ris;
        }
        public string grammiCad { get; set; }
        public string stecche { get; set; }
        public string nome { get; set; }

    }
}
