using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace StampaFax
{
    /// <summary>
    /// Logica di interazione per Window1.xaml
    /// </summary>
    public partial class FrmImpostaGiorno : Window
    {
        int giornoSettimana;
        XmlConfig xmlConfig;

        public FrmImpostaGiorno( XmlConfig xmlConfig)
        {
            InitializeComponent();

            giornoSettimana = xmlConfig.GiornoSettimanaPrelievo;
            this.xmlConfig = xmlConfig;

            switch (giornoSettimana)
            {
                case 1:
                    optLunedì.IsChecked = true;
                    break;
                case 2:
                    optMartedì.IsChecked = true;
                    break;
                case 3:
                    optMercoledì.IsChecked = true;
                    break;
                case 4:
                    optGiovedì.IsChecked = true;
                    break;
                case 5:
                    optVenerdì.IsChecked = true;
                    break;
                case 6:
                    optSabato.IsChecked = true;
                    break;
                default:
                    giornoSettimana = 1;
                    optLunedì.IsChecked = true;
                    break;
            }
        }

        private void cmdAnnulla_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void cmdConferma_Click(object sender, RoutedEventArgs e)
        {
            //salvo le nuove impostazioni
            if (optLunedì.IsChecked == true) giornoSettimana = 1;
            if (optMartedì.IsChecked == true) giornoSettimana = 2;
            if (optMercoledì.IsChecked == true) giornoSettimana = 3;
            if (optGiovedì.IsChecked == true) giornoSettimana = 4;
            if (optVenerdì.IsChecked == true) giornoSettimana = 5;
            if (optSabato.IsChecked == true) giornoSettimana = 6;

            xmlConfig.GiornoSettimanaPrelievo = giornoSettimana;

            this.Close();
        }

        
    }
}
