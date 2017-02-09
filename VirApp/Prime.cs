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
    class Prime
    {
        private int code_prime;
        private string designation_prime;
        private double montant_prime;

        static SqlConnection connect = new SqlConnection(WorkSpace.conString);
        string Query;
        static SqlCommand command = new SqlCommand();

        public Prime( string name, double montant)
        {
            
            this.designation_prime = name;
            this.montant_prime = montant;
        }
        public void Add_Prime()
        {
            try
            {
                BD con = new BD();
                con.seConnecter();
                int i = con.nbLigne("TypePrime", " WHERE  DésignationPrime = '" + designation_prime + "' ");
                if (i<1)
                {
                    int code = con.nbLigne("TypePrime", " ")+1;
                    Query = "INSERT INTO TypePrime ( CodePrime, DésignationPrime, MontantPrime,DateCreatType, CodeUser ) VALUES('" + code + "','" + designation_prime + "','" + montant_prime + "','"+DateTime.Today+"',01)";
                    SqlCommand cmd = new SqlCommand(Query, con.connextion());
                    cmd.ExecuteNonQuery();
                    MessageBox.Show(" réussite de l ' addition ");
                    con.seDeconnecter();
                }
                else
                {
                    MessageBox.Show(" Oups!! Cette Prime  existe deja ! \n" );
                }
            }
            catch (SqlException ex)
            {
                MessageBox.Show(" Une erreur a éte produite ! \n" + ex.Message);
            }
        }
    }
}
