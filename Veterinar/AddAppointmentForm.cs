using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
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
        private ApplicationContext db;
        private int _appointmentId;

        public AddAppointmentForm() : this(0) { }
        public AddAppointmentForm(int appointmentId)
        {
            InitializeComponent();
            _appointmentId = appointmentId;

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json");
            var config = builder.Build();
            string connectionString = config.GetConnectionString("DefaultConnection");

            var optionsBuilder = new DbContextOptionsBuilder<ApplicationContext>();
            // Отключаем проверку пулов и лишнее логирование для SQLite
            var options = optionsBuilder.UseSqlite(connectionString).Options;
            db = new ApplicationContext(options);

            // Отключаем автоматическое отслеживание изменений на уровне контекста формы,
            // чтобы EF вообще не тратил ресурсы на кэширование миллиона строк
            db.ChangeTracker.QueryTrackingBehavior = QueryTrackingBehavior.NoTracking;

            this.Load += AddAppointmentForm_Load;
            this.FormClosed += AddAppointmentForm_FormClosed;
        }

        private async void AddAppointmentForm_Load(object sender, EventArgs e)
        {
            //if (_appointmentId > 0)
            //{
            //    // Берем только ID вакцины, не выкачивая весь объект
            //    var vaccineId = await db.Vaccines
            //        .Where(x => x.Id == _appointmentId)
            //        .Select(x => x.Id)
            //        .FirstOrDefaultAsync();

            //    if (vaccineId > 0)
            //    {
            //        // Достаем ТОЛЬКО нужные поля в строго типизированный объект
            //        var currentItem = await db.Vaccines
            //            .Where(x => x.Id == vaccineId)
            //            .Select(x => new LookupItem { Id = (int)x.Id, Name = x.Name })
            //            .FirstOrDefaultAsync();

            //        if (currentItem != null)
            //        {
            //            // Оборачиваем в List<LookupItem>
            //            comboBox3.DataSource = new List<LookupItem> { currentItem };
            //            comboBox3.DisplayMember = "Name";
            //            comboBox3.ValueMember = "Id";
            //            comboBox3.SelectedValue = currentItem.Id;
            //        }
            //    }
            //}
        }
        public class LookupItem
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            // ВНИМАНИЕ: Убедитесь, что внутри VaccinesList() нет вызова .Load()!
            using (VaccinesList vaccinesList = new VaccinesList())
            {
                if (vaccinesList.ShowDialog(this) == DialogResult.Cancel) return;

                var vaccine = await db.Vaccines
                    .Where(v => v.Id == GlobalData.Data)
                    .Select(v => new LookupItem { Id = (int)v.Id, Name = v.Name })
                    .FirstOrDefaultAsync();

                if (vaccine != null)
                {
                    comboBox3.DataSource = new List<LookupItem> { vaccine };
                    comboBox3.DisplayMember = "Name";
                    comboBox3.ValueMember = "Id";
                    comboBox3.SelectedValue = vaccine.Id;
                }
            }
        }

        private async void button4_Click(object sender, EventArgs e)
        {
            // ВНИМАНИЕ: Убедитесь, что внутри PetsList() нет вызова .Load()!
            using (PetsList petsList = new PetsList())
            {
                if (petsList.ShowDialog(this) == DialogResult.Cancel) return;

                var pet = await db.Pets
                    .Where(p => p.Id == GlobalData.Data)
                    .Select(p => new LookupItem { Id = p.Id, Name = p.PetName }) // Исправлено: pet.PetName -> Name
                    .FirstOrDefaultAsync();

                if (pet != null)
                {
                    comboBox2.DataSource = new List<LookupItem> { pet };
                    comboBox2.DisplayMember = "Name";
                    comboBox2.ValueMember = "Id";
                    comboBox2.SelectedValue = pet.Id;
                }
            }
        }

        private void AddAppointmentForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            db?.Dispose();
        }
        private bool _isUpdatingDataSource = false;
        private async void comboBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_TextUpdate(object sender, EventArgs e)
        {
            string searchText = comboBox1.Text;

            // Начинаем поиск только если введено хотя бы 3 символа
            if (searchText.Length >= 3)
            {
                // Временно отключаем событие, чтобы обновление DataSource не вызвало повторный запуск этого же метода
                comboBox1.TextUpdate -= comboBox1_TextUpdate;

                // Используем ваше готовое поле db вместо создания нового контекста
                var filteredClients = db.Clients
                    .Where(c => c.Name.Contains(searchText))
                    .Take(20) // Строго ограничиваем выборку до 20 записей
                    .ToList();

                // Обновляем данные в ComboBox
                comboBox1.DataSource = filteredClients;
                comboBox1.ValueMember = "Id";
                comboBox1.DisplayMember = "Name";

                // Возвращаем текст пользователя, возвращаем курсор в конец строки и открываем выпадающий список
                comboBox1.Text = searchText;
                comboBox1.SelectionStart = searchText.Length;
                comboBox1.DroppedDown = true;

                // Включаем событие обратно
                comboBox1.TextUpdate += comboBox1_TextUpdate;
            }
        }
    }
}
