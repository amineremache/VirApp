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
    class BD
    {
        static SqlConnection maconnex = null;
        

        public void seConnecter()
        {
            maconnex = new SqlConnection(WorkSpace.conString);
            try
            {
                maconnex.Open();
            }
            catch (SqlException exep)
            {
                MessageBox.Show("Echec de connexion!" + exep.Message);
            }
        }
        // ---------------------------------------- seDeconnecter -----------------------------------------------//
        public void seDeconnecter()
        {
            try
            {
                maconnex.Close();

            }
            catch (SqlException exep)
            {
                MessageBox.Show("Echec de deconnexion! " + exep.Message);
            }
        }

        public SqlConnection connextion()
        {
            return maconnex;
        }
        public void executerRequete(String requete)
        {
            try
            {

                SqlCommand myCommand = new SqlCommand(requete, maconnex);
                myCommand.ExecuteNonQuery();
            }
            catch (SqlException e)
            {
                MessageBox.Show("Erreur !" + e.Message);
            }

        }

        public SqlDataReader getResultatRequete(String requete)
        {
            SqlDataReader resultat;
            try
            {
                SqlCommand myCommand = new SqlCommand(requete, maconnex);
                resultat = myCommand.ExecuteReader();

                return resultat;
            }
            catch (SqlException e)
            {
                MessageBox.Show("Erreur !" + e.Message);
                return null;
            }



        }

        public SqlDataAdapter getDataAdapter(string Query)
        {

            SqlCommand cmd = new SqlCommand(Query, maconnex);
            cmd.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            return da;
        }

        public string[] ID_colon = new string[8] { "Matricule", "NumDem", "CodeUser", "CodePV", "NumDem", "CodeVir", "CodeBanque", "CodePrime" };
        public string[] NOM_colone = new string[8] { "NomFonct", "", "NomUser", "", "", "", "DésignationBanque", "DésignationPrime" };
        public string[] PRENOM_colone = new string[8] { "PrenFonct", "", "PrenUser", "", "", "", "", "" };


        public void delete_query(string Table_Name, string condition_column, int condition, string instructor)
        {
            string query = "DELETE FROM " + Table_Name + " WHERE " + condition_column + "  " + instructor + "  '" + condition + "' ";
            executerRequete(query);

        }
        public void delete_query(string Table_Name, string condition_column, string condition, string instructor)
        {
            string query = "DELETE FROM " + Table_Name + " WHERE " + condition_column + "  " + instructor + "  '" + condition + "' ";
            executerRequete(query);

        }
        public string Search_Column(string nom_colone, string nom_tableau)
        {
            return ("SELECT  " + nom_colone + "  FROM " + nom_tableau + "");


        }
        public string Search_Where(string nom_colone, string nom_tableau, string condition_column, string condition)
        {
            return ("SELECT " + nom_colone + " FROM " + nom_tableau + " WHERE " + condition_column + " LIKE   '" + condition + "%'");

        }
        public string Search_Where(string nom_colone, string nom_tableau, string condition_column, long condition)
        {
            return ("SELECT " + nom_colone + " FROM " + nom_tableau + " WHERE " + condition_column + " LIKE   '" + condition + "%' ");

        }
        public string Search_Where(string nom_colone, string nom_tableau, string condition_column, DateTime condition)
        {
            return ("SELECT " + nom_colone + " FROM " + nom_tableau + " WHERE " + condition_column + " = '" + condition + "'");

        }
        public string Search_Where_instructer(string nom_colone, string nom_tableau, string condition_column, long condition, char instructor)
        {
            return ("SELECT " + nom_colone + "  FROM " + nom_tableau + "  WHERE " + condition_column + " " + instructor + " '" + condition + "'");

        }
        public string Search_Where_instructer(string nom_colone, string nom_tableau, string condition_column, string condition, char instructor)
        {
            return ("SELECT " + nom_colone + "  FROM " + nom_tableau + "  WHERE " + condition_column + " " + instructor + " '" + condition + "'");

        }
        public string update(string nom_tableau, string nom_colone, string valeur, string condition_colum, string instructor, string condition)
        {
            string update = "UPDATE " + nom_tableau + " SET " + nom_colone + " = '" + valeur + "' WHERE " + condition_colum + " " + instructor + "  '" + condition + "'";
            return update;

        }

        public string update(string nom_tableau, string nom_colone, string valeur, string condition_colum, string instructor, int condition)
        {
            string update = "UPDATE " + nom_tableau + " SET " + nom_colone + " = '" + valeur + "' WHERE " + condition_colum + "  " + instructor + "  " + condition + "";
            return update;

        }




        public int nbLigne(String table, String condition)
        {
            int cpt = 0;
            SqlDataReader resultat = getResultatRequete("SELECT count(*) FROM  " + table + " " + condition);
            try
            {
                while (resultat.Read())
                {
                    cpt = (int)resultat[0];

                }
                resultat.Close();
            }
            catch (SqlException e)
            {
                MessageBox.Show(e.Message);
            }
            return cpt;
        }
        public object[][] Rec_Fonc(string condition)
        {
            int nb_element = nbLigne("Fonctionnaire", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM Fonctionnaire  " + condition + " ");
            Object[][] fonc = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], a[9] };
                    fonc[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return fonc;
        }

       
        public object[][] Rec_Demande(string condition)
        {
            int nb_element = nbLigne("DemandePrime", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM DemandePrime  " + condition + " ");
            Object[][] demande = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2], a[3], a[4], a[5] };
                    demande[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return demande;

        }

        public object[][] Rec_EtatDemande(string condition)
        {
            int nb_element = nbLigne("TraitDemandes", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM TraitDemandes  " + condition + " ");
            Object[][] etatdemande = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2] };
                    etatdemande[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return etatdemande;

        }

        public object[][] Rec_Banque(string condition)
        {
            int nb_element = nbLigne("Banque", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM Banque  " + condition + " ");
            Object[][] banque = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2] };
                    banque[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return banque;

        }

        public object[][] Rec_PV(string condition)
        {
            int nb_element = nbLigne("PV", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM PV  " + condition + " ");
            Object[][] pv = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2] };
                    pv[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return pv;

        }

        public object[][] Rec_User(string condition)
        {
            int nb_element = nbLigne("utilisateur", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM utilisateur  " + condition + " ");
            Object[][] user = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2], a[3], a[4], a[5] };
                    user[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return user;

        }


        public object[][] Rec_Virement(string condition)
        {
            int nb_element = nbLigne("Virement", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM Virement  " + condition + " ");
            Object[][] Vir = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2] };
                    Vir[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return Vir;

        }


        public object[][] Rec_Prime(string condition)
        {
            int nb_element = nbLigne("TypePrime", condition);
            SqlDataReader a = getResultatRequete("SELECT * FROM TypePrime  " + condition + " ");
            Object[][] prime = new object[nb_element][];
            int cpt = 0;
            try
            {
                while (a.Read())
                {
                    object[] element = { a[0], a[1], a[2] };
                    prime[cpt] = element;
                    cpt++;
                }
            }
            catch (SqlException e)
            {
                MessageBox.Show("Error ! " + e);
            }

            return prime;

        }
    }
}
