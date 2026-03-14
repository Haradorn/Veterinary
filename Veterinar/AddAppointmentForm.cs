using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
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

namespace Veterinar
{
    public partial class AddAppointmentForm : Form
    {
        ApplicationContext db;
        public AddAppointmentForm()
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
        }

        private void button3_Click(object sender, EventArgs e)
        {
            VaccinesList vaccinesList = new VaccinesList();

            DialogResult result = vaccinesList.ShowDialog(this);
            if (result == DialogResult.Cancel) return;
            else
            {

                db.Vaccines.Load();
                Vaccine appointment = db.Vaccines.Find(GlobalData.Data);
                comboBox3.SelectedValue = appointment.Id;
            }

        }

        private void textBox1_Click(object sender, EventArgs e)
        {
        }
    }
}
