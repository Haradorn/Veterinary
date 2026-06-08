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
    public partial class PetsList : Form
    {
        ApplicationContext db;
        public PetsList()
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
            var pets = db.Pets.Include(o => o.Client).ToList();
            dataGridView1.DataSource = pets;
        }

        private void PetsList_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns["Id"].Visible = false;
            dataGridView1.Columns["ClientId"].Visible = false;
            dataGridView1.Columns["Breed"].HeaderText = "Порода";
            dataGridView1.Columns["Date"].HeaderText = "Дата рождения";
            dataGridView1.Columns["PetName"].HeaderText = "Кличка";
            dataGridView1.Columns["Client"].HeaderText = "Хозяин";
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
                Pet pet = db.Pets.Find(id);
                GlobalData.Data = pet.Id;
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
            var pet = db.Pets.Where(r => EF.Functions.Like(r.PetName, String.Format("%" + textBox1.Text + "%"))).ToList();
            dataGridView1.DataSource = pet;
            dataGridView1.Columns["Id"].Visible = false;
            dataGridView1.Columns["ClientId"].Visible = false;
            dataGridView1.Columns["Breed"].HeaderText = "Порода";
            dataGridView1.Columns["Date"].HeaderText = "Дата рождения";
            dataGridView1.Columns["PetName"].HeaderText = "Кличка";
            dataGridView1.Columns["Client"].HeaderText = "Хозяин";
        }
    }
}
