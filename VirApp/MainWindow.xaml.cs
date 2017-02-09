using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace VirApp
{

    public partial class MainWindow : MahApps.Metro.Controls.MetroWindow
    {
        public MainWindow()
        {

            InitializeComponent();



        }




        public static int CodeUser;
        public static string Droit;





        private void Aide_Click(object sender, RoutedEventArgs e)
        {
             System.Diagnostics.Process.Start(@""+WorkSpace.path + "\\aide_en_ligne\\templates\\admin\\help.html"); 
           
        }

    }

}
