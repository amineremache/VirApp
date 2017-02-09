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
using System.Drawing;
using System.Data.SqlClient;
using System.Data;

namespace VirApp
{
    class PV
    {
        private int code_PV;
        private DateTime date_PV;


        static SqlConnection connect = new SqlConnection(WorkSpace.conString);
        string Query;
        static SqlCommand command = new SqlCommand();



        public PV(int code)
        {
            this.code_PV = code;
            
        }

        public void Add_PV()
        {
            try
            {
               
                if (connect.State == ConnectionState.Closed) connect.Open();

                Query = "INSERT INTO PV ( CodePV, DatePV, CodeUser,Virement) VALUES(" + code_PV + ",'" + DateTime.Today.ToString("MM-dd-yyyy") + "',"+MainWindow.CodeUser+",1)";
                SqlCommand cmd = new SqlCommand(Query, connect);

                
                cmd.ExecuteNonQuery();
                MessageBox.Show("PV N°" + code_PV + " ajouté avec succèes !");

                if (connect.State == ConnectionState.Open) connect.Close();

                
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
