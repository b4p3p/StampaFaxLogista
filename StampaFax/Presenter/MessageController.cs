
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace StampaFax.Presenter
{
    class MessageController
    {
        internal static void Errore(string msg)
        {
            MessageBox.Show(msg, "Errore applicazione", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
