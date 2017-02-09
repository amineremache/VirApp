using Syncfusion.UI.Xaml.Charts;
using Syncfusion.Windows.Shared;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace VirApp
{

    public partial class Stats : MahApps.Metro.Controls.MetroWindow
    {
        public Stats()
        {
            InitializeComponent();
        }


        // PARTIE MVVM

        public Diagram DiagramUpdate;

        private void MenuItemAdv_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult resultat = MessageBox.Show("Êtes-vous sûre de vouloir quitter ?", "Quiter", MessageBoxButton.YesNo, MessageBoxImage.Question);

            //if (resultat == MessageBoxResult.Yes) Application.Current.Shutdown();
            if (resultat == MessageBoxResult.Yes) this.Close();

        }

        private void Reports_click1(object sender, RoutedEventArgs e)
        {
            gridStats.Visibility = Visibility.Visible;
            if (DiagramUpdate != null) DiagramUpdate.setState(Diagramme.chart);
            else sourceDiagram.setState(Diagramme.chart);
        }
        private void Reports_click2(object sender, RoutedEventArgs e)
        {
            gridStats.Visibility = Visibility.Visible;
            if (DiagramUpdate != null) DiagramUpdate.setState(Diagramme.dougnout);
            else sourceDiagram.setState(Diagramme.dougnout);
        }
        private void Reports_click3(object sender, RoutedEventArgs e)
        {
            gridStats.Visibility = Visibility.Visible;
            if (DiagramUpdate != null) DiagramUpdate.setState(Diagramme.histogram);
            else sourceDiagram.setState(Diagramme.histogram);
        }

        private void MenuItemAdv_Click_1(object sender, RoutedEventArgs e)
        {
            gridStats.Visibility = Visibility.Visible;
            if (DiagramUpdate != null) DiagramUpdate.setState(Diagramme.Don);
            else sourceDiagram.setState(Diagramme.Don);
        }


        private void button_Click(object sender, RoutedEventArgs e)
        {

            Diagram.DEBUT = date_debut;
            Diagram.FIN = date_fin;


            gridStats.Children.Clear();
            DiagramUpdate = new Diagram();

            gridStats.Children.Add(DiagramUpdate);
            MessageBox.Show("Les données ont été mis à jour ! ");



        }

    }
}
