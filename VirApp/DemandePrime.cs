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
    class DemandePrime
    {
        public void Add_DemandePrime(int num, DateTime date, int matri, int code, float montant, string compte, DateTime DateEv, string etatdem, string Motif, int codepv, DateTime Datecrea, int codeUse)
        {
            try
            {

                BD con = new BD();
                con.seConnecter();
                string txt1 = date.ToString("dd/mm/yyyy");
                string txt2 = DateEv.ToString("dd/mm/yyyy");
                string txt3 = Datecrea.ToString("dd/mm/yyyy");

                String requete = "INSERT INTO  DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontsantDem, CompteDem,DateEven,EtatDem,MotifEtat,pv_codepv,DateCreatDem,CodeUser) VALUES('" + num + "','" + txt1 + "','" + matri + "','" + code + "','" + montant + "','" + compte + "','" + txt2 + "','" + etatdem + "','" + Motif + "','" + codepv + "','" + txt3 + "','" + codeUse + "')";
                SqlCommand cmd = new SqlCommand(requete, con.connextion());
                cmd.ExecuteNonQuery();
                MessageBox.Show(" réussite de l ' addition ");
                con.seDeconnecter();
            }
            catch (SqlException ex)
            {
                MessageBox.Show(" Oups !Ce fonctionnaire existe deja ! \n" + ex.Message);
            }
        }

        public void Delete_DemandePrime(int numdem)
        {
            BD con = new BD();
            con.seConnecter();
            con.delete_query("DemandePrime", " NumDem", numdem, " = ");
            con.seDeconnecter();
            MessageBox.Show("réussite de la suppression");
        }

        internal void Add_DemandePrime()
        {
            throw new NotImplementedException();
        }
    }
}
