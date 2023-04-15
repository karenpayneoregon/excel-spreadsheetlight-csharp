using OleDbDemoForm.Classes;
using OleDbDemoForm.Extensions;

namespace OleDbDemoForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Example.Run("Somefile.xlsx");
            Shown += OnShown!;
        }

        private void OnShown(object sender, EventArgs e)
        {
            dataGridView1.DataSource = Operations.GetPeopleFromExcel();
            dataGridView1.FixHeaders();
            dataGridView1.ExpandColumns();
            dataGridView1.NoSort();
        }
        
    }
}
