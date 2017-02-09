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
    /// <summary>
    /// Interaction logic for Liste.xaml
    /// </summary>
    public partial class DataGridDialog : Window
    {
        public List<Employee> objList = new List<Employee>();
        private object dataGroupedGrid;
        private object dataNormalGrid;

        public DataGridDialog()
        {
            InitializeComponent();
            objList.Add(new Employee { EmpID = "Emp1", EmpName = "Yogi", EmpCity = "Chd" });
            objList.Add(new Employee { EmpID = "Emp2", EmpName = "Tom", EmpCity = "Chd" });
            objList.Add(new Employee { EmpID = "Emp3", EmpName = "Jerry", EmpCity = "Pkl" });
            objList.Add(new Employee { EmpID = "Emp4", EmpName = "Guffy", EmpCity = "Chd" });
            objList.Add(new Employee { EmpID = "Emp5", EmpName = "Donald", EmpCity = "Pkl" });
            objList.Add(new Employee { EmpID = "Emp6", EmpName = "Mickey", EmpCity = "Pkl" });
            objList.Add(new Employee { EmpID = "Emp7", EmpName = "Minny", EmpCity = "Chd" });
            objList.Add(new Employee { EmpID = "Emp8", EmpName = "Daffy", EmpCity = "Del" });
            objList.Add(new Employee { EmpID = "Emp9", EmpName = "LoadRunner", EmpCity = "Del" });
            objList.Add(new Employee { EmpID = "Emp10", EmpName = "Bluto", EmpCity = "Del" });
            //dataNormalGrid.ItemsSource = objList;
            var collectionVwSrc = new ListCollectionView(objList);
            collectionVwSrc.GroupDescriptions.Add(new PropertyGroupDescription("EmpCity"));
            //dataGroupedGrid.ItemsSource = collectionVwSrc;
        }

        private void InitializeComponent()
        {
            throw new NotImplementedException();
        }
    }
    public class Employee
    {
        public string EmpID { get; set; }
        public string EmpName { get; set; }
        public string EmpCity { get; set; }
    }
}
