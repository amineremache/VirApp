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
    class cotis
    {
        
        // cette methode concerne une demande effectuer par un seul fonctionnaire 
        //elle verifie si les conditions sont vérifier et ajoute une demande dqans la table des demandes

        public void acc_ref_demandes(int matricule, DateTime date_event, String designation, int codeuser, int codepv, DateTime datedem, DateTime datecreadem)
        {

            //se connecter à la base de données
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();

            //vérifier si le fonctionnaire existe dans la base de données
            SqlCommand cmd = new SqlCommand("SELECT * FROM Fonctionnaire WHERE Matricule='" + matricule + "' ", con);
            SqlDataReader read = cmd.ExecuteReader();
            con.Close();

            if (read.Read()) //si le fonctionnaire existe
            {


                DateTime date = Convert.ToDateTime(read[3]);

                if (date.Date > date_event.Date) //vérifier si le fonctiooanire est recruter avant l'événement 
                { MessageBox.Show("Vous etes recrutez apres cette evenement"); }
                else//s'il la condition est vérifiée
                {

                    con.Open();
                    int year = Convert.ToDateTime(date_event).Year;
                    int mois = Convert.ToDateTime(date_event).Month;
                    int jour = Convert.ToDateTime(date_event).Day;

                    //verifier si la demande a ete faite avant
                    cmd = new SqlCommand("SELECT * FROM  WHERE Matricule=" + matricule + "AND  DATEDIFF(DAY,'" + year + "-" + mois + "-" + jour + "',DateEven)=0 ", con);
                    read = cmd.ExecuteReader();
                    con.Close();

                    if (read.Read()) //si elle est faite on doit la mentionner
                    {
                        string etat = Convert.ToString(read[7]);
                        MessageBox.Show("vous avez depose la demande et votre etat etait:" + etat);

                    }
                    else //si la demande n'existe pas on va l'ajouter
                    {

                        DemandePrime dem = new DemandePrime();
                        //compter le nombre de demandes
                        con.Open();
                        cmd = new SqlCommand("SELECT * FROM DemandePrime ", con);
                        read = cmd.ExecuteReader();
                        int cpt = 1;
                        while (read.Read())
                        { cpt++; }
                        con.Close();
                        //avoir le montant de la prime et le code de la prime
                        con.Open();
                        cmd = new SqlCommand("SELECT * FROM TypePrime WHERE   DésignationPrime= " + designation + "", con);
                        read = cmd.ExecuteReader();
                        read.Read();
                        //Avoir les informations du fonctionnaire
                        cmd = new SqlCommand("SELECT * FROM Fonctionnaire WHERE  Matricule = " + matricule + "", con);
                        SqlDataReader read1 = cmd.ExecuteReader();
                        read1.Read();
                        // Avoir le code de l'utilisateur (c'est un  parametre d'entrée)

                        //On ajoute la demande dans la table des 
                       
                        dem.Add_DemandePrime(cpt, datedem, matricule, Convert.ToInt32(read[0]), Convert.ToSingle(read[2]), Convert.ToString(read1[6]), date_event, "N", null, codepv, datecreadem, codeuser);
                    }

                }

            }

            else // si le fonctionnaire existe on doit montionner ça
            { MessageBox.Show("Le fonctionnaire n'existe pas"); }

            con.Close(); // fermer la base de données

        }

        //********************************************************************************************************************//

        // cette methode genere une liste des fonctionnaires concerner par batch (acceptes) et les mettre dans un data table 

        public SqlDataAdapter batch_accepte()
        {


            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM Parametres ", con);
            SqlDataReader read = cmd.ExecuteReader();
            read.Read();
            int durcot = Convert.ToInt32(read[2]);
            int jourd = Convert.ToInt32(read[3]);
            int moideb = Convert.ToInt32(read[4]);
            int year = Convert.ToInt32(DateTime.Now.Year);
            int month = Convert.ToInt32(DateTime.Now.Month);
            con.Close();

            if (month >= moideb && month <= 12) { }
            else { year--; }

            // La requete qui selectionne tout les fonctioonaires qui sont acceptes dans la prime generale
            con.Open();
            SqlDataAdapter ad = new SqlDataAdapter("SELECT * FROM Fonctionnaire WHERE (DateDepartTmp IS NULL AND DateDepartDefi IS NULL) OR(   DateDepartDefi IS  NOT NULL AND  DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartDefi)>=" + durcot + ")OR( DateDepartTmp IS NOT NULL AND DateRetrTmp IS NULL AND DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartTmp)>=" + durcot + ")OR(NOT( DateDepartTmp IS NULL)  AND   (NOT DateRetrTmp IS NULL ) )", con);


            con.Close();


            // retourner notre data table 
            return ad;
        }



        //************************************************************************************************************************//

        //une methode qui retourne un data table qui contient les fonctionnaires non concernes par les primes generale
        public SqlDataAdapter batch_Non()
        {


            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM Parametres ", con);
            SqlDataReader read = cmd.ExecuteReader();
            read.Read();
            int durcot = Convert.ToInt32(read[2]);
            int jourd = Convert.ToInt32(read[3]);
            int moideb = Convert.ToInt32(read[4]);
            int year = Convert.ToInt32(DateTime.Now.Year);
            int month = Convert.ToInt32(DateTime.Now.Month);
            con.Close();
            if (month >= moideb && month <= 12) { }
            else { year--; }
            con.Open();

            // La requete qui selectionne tout les fonctioonaires qui sont acceptes dans la prime generale
            SqlDataAdapter ad = new SqlDataAdapter("SELECT * FROM Fonctionnaire WHERE NOT( (DateDepartTmp IS NULL AND DateDepartDefi IS NULL) OR(   DateDepartDefi IS  NOT NULL AND  DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartDefi)>=" + durcot + ")OR( DateDepartTmp IS NOT NULL AND DateRetrTmp IS NULL AND DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartTmp)>=" + durcot + ")OR   ( DateDepartTmp IS NOT NULL AND DateRetrTmp IS NOT NULL ) )", con);


            con.Close(); //fermer la base de données


            // retourner notre data table 
            return ad;
        }

        //**********************************************************************************************************************//

        // modifier la durée de cotisation

        public void modifier_dure(int dure)
        {
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE Parametres SET DurCot='" + dure + "'", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        //modifier le jour et le mois du debut de l'année sociale 
        public void modif_jourMois(int jour, int mois)
        {
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE Parametres SET JourDebAnSoc='" + jour + "', MoisDebAnSoc='" + mois + "'", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        //*******************************************************************************************************************************//

        //Marquer un depart Tmp 

        public void marque_deperTmp(DateTime Date, String motif)
        {
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            SqlCommand cmd;
            con.Open();
            cmd = new SqlCommand("UPDATE Fonctionnaires SET DateDepartTmp='" + Date + "', MotifDepartTmp ='" + motif + "' ", con);
            cmd.ExecuteNonQuery();

            con.Close();
        }
        //marquer un retour Tmp

        public void marque_retourTmp(DateTime Date, String motif)
        {
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE Fonctionnaires SET DateRetrTmp='" + Date + "', MotifDepartTmp ='" + motif + "' ", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        //marquer un depart Def
        public void marque_departDef(DateTime Date, String motif)
        {
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("UPDATE Fonctionnaires SET DateDepartDefi='" + Date + "', MotifDepartDefi ='" + motif + "' ", con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        //***************************************************************************************************************************************//

        // Une methode pour Ajouter la liste de batch à la table des demandes
        public void ajouter_liste(SqlDataReader liste, String typeprime, string datedem, int codeuser, int codepv)
        {
            //connecter la base de données
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            SqlCommand cmd;
            SqlDataReader read;
            DemandePrime dem = new DemandePrime();


            //compter le nombre de demande dans la table des demandes

            


            string key="0";

            
            if (WorkSpace.GetterLastKey() != "")
            {
                key = ((Int32.Parse(WorkSpace.GetterLastKey()) / 100) ).ToString() + Int32.Parse(DateTime.Today.ToString("yy"));
                if (key.Length == 3) key = "00" + key;
                else if (key.Length == 4) key = "0" + key;
            }

            else key = "0";

            
             int cle = int.Parse(key)/100 *100 + int.Parse(DateTime.Today.ToString("yy"));

            
            //Tirer la prime de la table des type de prime pour avoir son code et sa valeur
            //****************************************//
            con.Open();
            cmd = new SqlCommand("SELECT * FROM TypePrime WHERE   DésignationPrime= '" + typeprime + "'", con);
            read = cmd.ExecuteReader();
            read.Read();
            int code = (int)read[0];
            float montant = Convert.ToSingle(read[2]);
            con.Close();
            //commancer a ajouter les demandes dans la table des demandes
            con.Open();
            //************************************//
            while (liste.Read())
            {


                try
                {
                    cle += 100;

                    

                    int jour = int.Parse(datedem.Substring(0,2));
                    int mois = int.Parse(datedem.Substring(3,2));
                    int an = int.Parse(datedem.Substring(6,4));

                    
                    
                    string requete = "INSERT INTO  DemandePrime (NumDem, DateDem, Matricule, CodePrime, MontantDem, CompteDem,DateEven,EtatDem,MotifEtat,pv_codepv,DateCreatDem,CodeUser) VALUES (" + cle + ", '" + an + "-" + mois + "-" + jour + "', '" + Convert.ToString(liste[0]) + "'," + code + "," + montant + ", '" + Convert.ToString(liste[6]) + "', '" + an + "-" + mois + "-" + jour + "', 'A', null ,"+ codepv +", '" + an + "-" + mois + "-" + jour + "', " + codeuser + ")";
                    cmd = new SqlCommand(requete, con);
                    cmd.ExecuteNonQuery();
                    
                    
                    WorkSpace.lastKey = cle.ToString();
                    



                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
            con.Close();

        }
        //**********************************************************************************************************************************//

        //une méthode qui retourne unn datareader dela liste des fonctionnaires qui méritent la prime 
        // on a besion de cette méthode pour ajouter les prime dans la atbble des demandes 
        public SqlDataReader batch_acc()
        {


            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM Parametres ", con);
            SqlDataReader read = cmd.ExecuteReader();
            read.Read();
            int durcot = Convert.ToInt32(read[2]);
            int jourd = Convert.ToInt32(read[3]);
            int moideb = Convert.ToInt32(read[4]);
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            con.Close();
            int M = Convert.ToInt32(month);
            int Y = Convert.ToInt32(year);
            if (M >= moideb && M <= 12) { }
            else { Y--; }

            // La requete qui selectionne tout les fonctioonaires qui sont acceptes dans la prime generale
            con.Open();
            cmd = new SqlCommand("SELECT * FROM Fonctionnaire WHERE (DateDepartTmp IS NULL AND DateDepartDefi IS NULL) OR(   DateDepartDefi IS  NOT NULL AND  DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartDefi)>=" + durcot + ")OR( DateDepartTmp IS NOT NULL AND DateRetrTmp IS NULL AND DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartTmp)>=" + durcot + ")OR(NOT( DateDepartTmp IS NULL)  AND   (NOT DateRetrTmp IS NULL ) )", con);
            SqlDataReader lire = cmd.ExecuteReader();


            // retourner notre datareader
            return lire;
            con.Close();

        }

        // les employées refusé pour le batch

        public SqlDataReader batch_ref()
        {
            SqlConnection con = new SqlConnection(WorkSpace.conString);
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT * FROM Parametres ", con);
            SqlDataReader read = cmd.ExecuteReader();
            read.Read();
            int durcot = Convert.ToInt32(read[2]);
            int jourd = Convert.ToInt32(read[3]);
            int moideb = Convert.ToInt32(read[4]);
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;
            con.Close();
            int M = Convert.ToInt32(month);
            int Y = Convert.ToInt32(year);
            if (M >= moideb && M <= 12) { }
            else { Y--; }

            // La requete qui selectionne tout les fonctioonaires qui sont acceptes dans la prime generale
            con.Open();
            cmd = new SqlCommand("SELECT * FROM Fonctionnaire WHERE NOT( (DateDepartTmp IS NULL AND DateDepartDefi IS NULL) OR(   DateDepartDefi IS  NOT NULL AND  DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartDefi)>=" + durcot + ")OR( DateDepartTmp IS NOT NULL AND DateRetrTmp IS NULL AND DATEDIFF(DAY,'" + year + "-" + moideb + "-" + jourd + "',DateDepartTmp)>=" + durcot + ")OR(NOT( DateDepartTmp IS NULL)  AND   (NOT DateRetrTmp IS NULL ) ))", con);
            SqlDataReader lire = cmd.ExecuteReader();


            // retourner notre datareader
            return lire;
            con.Close();




        }

    }
}
