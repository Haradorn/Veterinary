using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Veterinar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static Veterinar.VaccinesList.CacheSimulator;

namespace Veterinar
{
    public partial class VaccinesList : Form
    {
        ApplicationContext db;

        private const int TotalRows = 1000000;

        public VaccinesList()
        {
            InitializeComponent();
            SetupGrid();
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

        private void SetupGrid()
        {
            // 1. Включаем виртуальный режим (КРИТИЧЕСКИ ВАЖНО)
            dataGridView1.VirtualMode = true;

            // 2. Отключаем автоматический расчет размеров (иначе зависнет)
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            // 3. Базовые настройки интерфейса
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.ReadOnly = true; // В данном примере только для чтения

            // 4. Вручную создаем колонки
            dataGridView1.Columns.Add("Id", "ID");
            dataGridView1.Columns.Add("Name", "Название вакцины");

            // Устанавливаем фиксированную ширину колонок
            dataGridView1.Columns[0].Width = 80;
            dataGridView1.Columns[1].Width = 200;

            // 5. Главный шаг: подписываемся на событие запроса данных ячейки
            dataGridView1.CellValueNeeded += DataGridView1_CellValueNeeded;

            // 6. Выделяем память под виртуальные строки (создается только скроллбар)
            dataGridView1.RowCount = TotalRows;
        }

        private void DataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= TotalRows) return;

            // Получаем данные для всей строки из нашего SQLite-кэша
            CacheSimulator.VaccineData rowData = CacheSimulator.GetRowData(e.RowIndex);

            switch (e.ColumnIndex)
            {
                case 0:
                    e.Value = e.RowIndex + 1; // Порядковый номер в таблице
                    break;
                case 1:
                    e.Value = rowData.Name;   // Из кэша
                    break;
            }
        }

        public static class CacheSimulator
        {
            // Путь к вашей базе данных SQLite
            private static readonly string ConnectionString = "Data Source=veterinar.db;Version=3;";

            // Размер одной страницы (пачки) данных для загрузки из БД
            private const int PageSize = 100;

            // Номер страницы, которая сейчас хранится в оперативной памяти
            private static int _currentPageNumber = -1;

            // Локальный кэш текущей страницы в памяти (хранит ID строки и объект данных)
            private static readonly Dictionary<int, VaccineData> _pageCache = new Dictionary<int, VaccineData>();

            // Структура для хранения данных одной строки (занимает минимум памяти)
            public struct VaccineData
            {
                public string Name;
            }

            /// <summary>
            /// Главный метод, который вызывает DataGridView через CellValueNeeded
            /// </summary>
            public static VaccineData GetRowData(int rowIndex)
            {
                // Округляем индекс строки, чтобы узнать, к какой «странице» она относится
                int targetPage = rowIndex / PageSize;

                // Если нужной страницы нет в памяти — загружаем её из SQLite
                if (targetPage != _currentPageNumber)
                {
                    LoadPageFromDatabase(targetPage);
                }

                // Возвращаем данные из быстрого кэша в памяти
                if (_pageCache.TryGetValue(rowIndex, out var data))
                {
                    return data;
                }

                // Заглушка на случай, если строка в базе отсутствует
                return new VaccineData { Name = "Ошибка"};
            }

            /// <summary>
            /// Загрузка пачки строк из SQLite в оперативную память
            /// </summary>
            private static void LoadPageFromDatabase(int pageNumber)
            {
                _pageCache.Clear(); // Очищаем старый кэш, чтобы не забивать RAM
                _currentPageNumber = pageNumber;

                int offset = pageNumber * PageSize;

                using (var connection = new SQLiteConnection(ConnectionString))
                {
                    connection.Open();

                    // Запрос вытягивает строго PageSize (100) строк, начиная с нужного смещения.
                    // Использование LIMIT/OFFSET эффективно, так как мы запрашиваем маленькие порции.
                    string query = "SELECT Name FROM Vaccines ORDER BY id LIMIT @limit OFFSET @offset";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@limit", PageSize);
                        command.Parameters.AddWithValue("@offset", offset);

                        using (var reader = command.ExecuteReader())
                        {
                            int currentId = offset;
                            while (reader.Read())
                            {
                                var vaccine = new VaccineData
                                {
                                    Name = reader.GetString(0),
                                };

                                _pageCache[currentId] = vaccine;
                                currentId++;
                            }
                        }
                    }
                }
            }
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
