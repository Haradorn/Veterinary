using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Veterinar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Veterinar
{
    public partial class VaccinesList : Form
    {
        ApplicationContext db;
        public VaccinesList()
        {
            InitializeComponent();
            var builder = new ConfigurationBuilder();
            // установка пути к текущему каталогу
            builder.SetBasePath(Directory.GetCurrentDirectory());
            // получаем конфигурацию из файла appsettings.json
            builder.AddJsonFile("appsettings.json");
            // создаем конфигурацию
            var config = builder.Build();
            // получаем строку подключения
            string connectionString = config.GetConnectionString("DefaultConnection");

            var optionsBuilder = new DbContextOptionsBuilder<ApplicationContext>();
            var options = optionsBuilder.UseSqlite(connectionString).Options;
            db = new ApplicationContext(options);
            db.Vaccines.Load();
            dataGridView1.DataSource = db.Vaccines.Local.ToBindingList();
            dataGridView1.Show();
        }

        private void VaccinesList_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns["Id"].Visible = false;
            dataGridView1.Columns["Name"].HeaderText = "Название вакцины";
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            int index = dataGridView1.SelectedRows[0].Index;
            int id = 0;
            bool converted = Int32.TryParse(dataGridView1[0, index].Value.ToString(), out id);
            if (converted == false)
                return;
            try
            {
                Vaccine vaccine = db.Vaccines.Find(id);
                GlobalData.Data = vaccine.Id;
                this.DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var vaccine = db.Vaccines.Where(r => EF.Functions.Like(r.Name, String.Format("%" + textBox1.Text + "%"))).ToList();
            dataGridView1.DataSource = vaccine;
            dataGridView1.Columns["Id"].Visible = false;
            dataGridView1.Columns["Name"].HeaderText = "Название вакцины";
        }
    }
}
