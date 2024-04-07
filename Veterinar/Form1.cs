using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Veterinar;
using System.IO;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace Veterinar
{
    public partial class Form1 : Form
    {
        ApplicationContext db;
        public Form1()
        {
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
            InitializeComponent();
            db.Appointments.Load();
            db.Clients.Load();
            db.Pets.Load();
            db.Vaccines.Load();
            dataGridView1.DataSource = db.Appointments.Local.ToBindingList();
            dataGridView2.DataSource = db.Clients.Local.ToBindingList();
            dataGridView3.DataSource = db.Pets.Local.ToBindingList();
            dataGridView4.DataSource = db.Vaccines.Local.ToBindingList();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddAppointmentForm appointmentForm = new AddAppointmentForm();
            List<Client> clients = db.Clients.ToList();
            appointmentForm.comboBox1.DataSource = clients;
            appointmentForm.comboBox1.ValueMember = "Id";
            appointmentForm.comboBox1.DisplayMember = "Name";

            List<Pet> pets = db.Pets.ToList();
            appointmentForm.comboBox2.DataSource = pets;
            appointmentForm.comboBox2.ValueMember = "Id";
            appointmentForm.comboBox2.DisplayMember = "PetName";

            List<Vaccine> vaccines = db.Vaccines.ToList();
            appointmentForm.comboBox3.DataSource = vaccines;
            appointmentForm.comboBox3.ValueMember = "Id";
            appointmentForm.comboBox3.DisplayMember = "Name";

            DialogResult result = appointmentForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            else
            {
                try
                {
                    Appointment appointment = new Appointment();
                    appointment.Date = appointmentForm.dateTimePicker1.Value;
                    appointment.Client = (Client)appointmentForm.comboBox1.SelectedItem;
                    appointment.Pet = (Pet)appointmentForm.comboBox2.SelectedItem;
                    appointment.Vaccine = (Vaccine)appointmentForm.comboBox3.SelectedItem;
                    appointment.Service = appointmentForm.richTextBox1.Text;
                    appointment.WhatHurt = appointmentForm.richTextBox2.Text;
                    appointment.WhatNeedToDo = appointmentForm.richTextBox3.Text;
                    appointment.WhatWasDone = appointmentForm.richTextBox4.Text;

                    db.Appointments.Add(appointment);
                    db.SaveChanges();
                    MessageBox.Show("Запись о приёме успешно добавлена");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    int index = dataGridView1.SelectedRows[0].Index;

                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView1[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Appointment appointment = db.Appointments.Find(id);
                    AddAppointmentForm addAppointment = new AddAppointmentForm();
                    List<Client> clients = db.Clients.ToList();
                    addAppointment.comboBox1.DataSource = clients;
                    addAppointment.comboBox1.ValueMember = "Id";
                    addAppointment.comboBox1.DisplayMember = "Name";
                    if (appointment.Client != null)
                        addAppointment.comboBox1.SelectedValue = appointment.Client.Id;
                    List<Pet> pets = db.Pets.ToList();
                    addAppointment.comboBox2.DataSource = pets;
                    addAppointment.comboBox2.ValueMember = "Id";
                    addAppointment.comboBox2.DisplayMember = "PetName";
                    if (appointment.Pet != null)
                        addAppointment.comboBox2.SelectedValue = appointment.Pet.Id;
                    List<Vaccine> vaccines = db.Vaccines.ToList();
                    addAppointment.comboBox3.DataSource = vaccines;
                    addAppointment.comboBox3.ValueMember = "Id";
                    addAppointment.comboBox3.DisplayMember = "Name";
                    if (appointment.Vaccine != null)
                        addAppointment.comboBox3.SelectedValue = appointment.Vaccine.Id;
                    addAppointment.richTextBox1.Text = appointment.Service;
                    addAppointment.richTextBox2.Text = appointment.WhatHurt;
                    addAppointment.richTextBox3.Text = appointment.WhatNeedToDo;
                    addAppointment.richTextBox4.Text = appointment.WhatWasDone;
                    addAppointment.dateTimePicker1.Value = appointment.Date;


                    DialogResult result = addAppointment.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    else
                    {

                        appointment.Client = (Client)addAppointment.comboBox1.SelectedItem;
                        appointment.Pet = (Pet)addAppointment.comboBox2.SelectedItem;
                        appointment.Vaccine = (Vaccine)addAppointment.comboBox3.SelectedItem;
                        appointment.Service = addAppointment.richTextBox1.Text;
                        appointment.WhatHurt = addAppointment.richTextBox2.Text;
                        appointment.WhatNeedToDo = addAppointment.richTextBox3.Text;
                        appointment.WhatWasDone = addAppointment.richTextBox4.Text;
                        appointment.Date = Convert.ToDateTime(addAppointment.dateTimePicker1.Value);

                        db.Entry(appointment).State = EntityState.Modified;
                        db.SaveChanges();
                        dataGridView1.Refresh();
                        MessageBox.Show("Запись о приёме обновлена");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                try
                {
                    int index = dataGridView1.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView1[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Appointment appointment = db.Appointments.Find(id);
                    db.Appointments.Remove(appointment);
                    db.SaveChanges();
                    MessageBox.Show("Запись о приёме удалена");
                    dataGridView1.Refresh();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns["Id"].Visible = false;
            dataGridView1.Columns["ClientId"].Visible = false;
            dataGridView1.Columns["PetId"].Visible = false;
            dataGridView1.Columns["VaccineId"].Visible = false;
            dataGridView1.Columns["Client"].HeaderText = "Клиент";
            dataGridView1.Columns["Pet"].HeaderText = "Питомец";
            dataGridView1.Columns["Date"].HeaderText = "Дата";
            dataGridView1.Columns["Vaccine"].HeaderText = "Вакцина";
            dataGridView1.Columns["Service"].HeaderText = "Услуга";
            dataGridView1.Columns["WhatHurt"].HeaderText = "Что болело";
            dataGridView1.Columns["WhatNeedToDo"].HeaderText = "Что нужно сделать";
            dataGridView1.Columns["WhatWasDone"].HeaderText = "Что было сделано";
            dataGridView2.Columns["Id"].Visible = false;
            dataGridView2.Columns["Address"].HeaderText = "Адрес";
            dataGridView2.Columns["Name"].HeaderText = "Имя клиента";
            dataGridView2.Columns["Phone"].HeaderText = "Телефон";
            dataGridView3.Columns["Id"].Visible = false;
            dataGridView3.Columns["ClientId"].Visible = false;
            dataGridView3.Columns["Breed"].HeaderText = "Порода";
            dataGridView3.Columns["Date"].HeaderText = "Дата рождения";
            dataGridView3.Columns["PetName"].HeaderText = "Кличка";
            dataGridView3.Columns["Client"].HeaderText = "Хозяин";
            dataGridView4.Columns["Id"].Visible = false;
            dataGridView4.Columns["Name"].HeaderText = "Название вакцины";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            AddClientForm addClientForm = new AddClientForm();

            DialogResult result = addClientForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            else
            {
                try
                {
                    Client client = new Client();
                    client.Name = addClientForm.textBox1.Text;
                    client.Phone = addClientForm.textBox2.Text;
                    client.Address = addClientForm.textBox3.Text;

                    db.Clients.Add(client);
                    db.SaveChanges();
                    dataGridView2.Update();
                    MessageBox.Show("Запись о клиенте успешно добавлена");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    int index = dataGridView2.SelectedRows[0].Index;

                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView2[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Client client = db.Clients.Find(id);
                    AddClientForm addClientForm = new AddClientForm();
                    addClientForm.textBox1.Text = client.Name;
                    addClientForm.textBox2.Text = client.Phone;
                    addClientForm.textBox3.Text = client.Address;

                    DialogResult result = addClientForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    else
                    {
                        client.Name = addClientForm.textBox1.Text;
                        client.Phone = addClientForm.textBox2.Text;
                        client.Address = addClientForm.textBox3.Text;


                        db.Entry(client).State = EntityState.Modified;
                        db.SaveChanges();
                        dataGridView2.Refresh();
                        MessageBox.Show("Запись о клиенте обновлена");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                try
                {
                    int index = dataGridView2.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView2[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Client client = db.Clients.Find(id);
                    db.Clients.Remove(client);
                    db.SaveChanges();
                    dataGridView2.Refresh();
                    MessageBox.Show("Запись о клиенте удалена");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            AddPetForm addPetForm = new AddPetForm();
            List<Client> clients = db.Clients.ToList();
            addPetForm.comboBox1.DataSource = clients;
            addPetForm.comboBox1.ValueMember = "Id";
            addPetForm.comboBox1.DisplayMember = "Name";

            DialogResult result = addPetForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            else
            {
                try
                {
                    Pet pet= new Pet();
                    pet.Breed = addPetForm.textBox1.Text;
                    pet.PetName = addPetForm.textBox2.Text;
                    pet.Client = (Client)addPetForm.comboBox1.SelectedItem;
                    pet.Date = addPetForm.dateTimePicker1.Value;

                    db.Pets.Add(pet);
                    db.SaveChanges();
                    MessageBox.Show("Запись о питомце добавлена");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView3.SelectedRows.Count > 0)
                {
                    int index = dataGridView3.SelectedRows[0].Index;

                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView3[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Pet pet = db.Pets.Find(id);
                    AddPetForm addPetForm = new AddPetForm();
                    List<Client> clients = db.Clients.ToList();
                    addPetForm.comboBox1.DataSource = clients;
                    addPetForm.comboBox1.ValueMember = "Id";
                    addPetForm.comboBox1.DisplayMember = "Name";
                    if (pet.Client != null)
                        addPetForm.comboBox1.SelectedValue = pet.Client.Id;
                    addPetForm.textBox1.Text = pet.Breed;
                    addPetForm.textBox2.Text = pet.PetName;
                    addPetForm.dateTimePicker1.Value = pet.Date;
                    DialogResult result = addPetForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    else
                    {
                        pet.Client = (Client)addPetForm.comboBox1.SelectedItem;
                        pet.Breed = addPetForm.textBox1.Text;
                        pet.PetName = addPetForm.textBox2.Text;
                        pet.Date = Convert.ToDateTime(addPetForm.dateTimePicker1.Value);

                        db.Entry(pet).State = EntityState.Modified;
                        db.SaveChanges();
                        dataGridView3.Refresh();
                        MessageBox.Show("Запись о питомце обновлена");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (dataGridView3.SelectedRows.Count > 0)
            {
                try
                {
                    int index = dataGridView3.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView3[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Pet pet = db.Pets.Find(id);
                    db.Pets.Remove(pet);
                    db.SaveChanges();
                    dataGridView3.Refresh();
                    MessageBox.Show("Запись о питомце удалена");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            AddVaccineForm addVaccineForm = new AddVaccineForm();

            DialogResult result = addVaccineForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            else
            {
                try
                {
                    Vaccine vaccine = new Vaccine();
                    vaccine.Name = addVaccineForm.textBox1.Text;

                    db.Vaccines.Add(vaccine);
                    db.SaveChanges();
                    MessageBox.Show("Запись о вакцине успешно добавлена");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView4.SelectedRows.Count > 0)
                {
                    int index = dataGridView4.SelectedRows[0].Index;

                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView4[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Vaccine vaccine = db.Vaccines.Find(id);
                    AddVaccineForm addVaccineForm = new AddVaccineForm();
                    addVaccineForm.textBox1.Text = vaccine.Name;
                    DialogResult result = addVaccineForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    else
                    {
                        vaccine.Name = addVaccineForm.textBox1.Text;
                        db.Entry(vaccine).State = EntityState.Modified;
                        db.SaveChanges();
                        dataGridView4.Refresh();
                        MessageBox.Show("Запись о вакцине обновлена");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (dataGridView4.SelectedRows.Count > 0)
            {
                try
                {
                    int index = dataGridView4.SelectedRows[0].Index;
                    int id = 0;
                    bool converted = Int32.TryParse(dataGridView4[0, index].Value.ToString(), out id);
                    if (converted == false)
                        return;
                    Vaccine vaccine = db.Vaccines.Find(id);
                    db.Vaccines.Remove(vaccine);
                    db.SaveChanges();
                    dataGridView4.Refresh();
                    MessageBox.Show("Запись о вакцине удалена");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            var appointment = db.Appointments.Where(r => EF.Functions.Like(r.Client.Name, String.Format("%" + textBox1.Text + "%"))).ToList();

            dataGridView1.DataSource = appointment;
            dataGridView1.Columns["Id"].Visible = false;
            dataGridView1.Columns["ClientId"].Visible = false;
            dataGridView1.Columns["PetId"].Visible = false;
            dataGridView1.Columns["VaccineId"].Visible = false;
            dataGridView1.Columns["Client"].HeaderText = "Клиент";
            dataGridView1.Columns["Pet"].HeaderText = "Питомец";
            dataGridView1.Columns["Date"].HeaderText = "Дата";
            dataGridView1.Columns["Vaccine"].HeaderText = "Вакцина";
            dataGridView1.Columns["Service"].HeaderText = "Услуга";
            dataGridView1.Columns["WhatHurt"].HeaderText = "Что болело";
            dataGridView1.Columns["WhatNeedToDo"].HeaderText = "Что нужно сделать";
            dataGridView1.Columns["WhatWasDone"].HeaderText = "Что было сделано";
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            var client = db.Clients.Where(r => EF.Functions.Like(r.Name, String.Format("%" + textBox2.Text + "%"))).ToList();
            dataGridView2.DataSource = client;
            dataGridView2.Columns["Id"].Visible = false;
            dataGridView2.Columns["Address"].HeaderText = "Адрес";
            dataGridView2.Columns["Name"].HeaderText = "Имя клиента";
            dataGridView2.Columns["Phone"].HeaderText = "Телефон";
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            var pet = db.Pets.Where(r => EF.Functions.Like(r.PetName, String.Format("%" + textBox3.Text + "%"))).ToList();
            dataGridView3.DataSource = pet;
            dataGridView3.Columns["Id"].Visible = false;
            dataGridView3.Columns["ClientId"].Visible = false;
            dataGridView3.Columns["Breed"].HeaderText = "Порода";
            dataGridView3.Columns["Date"].HeaderText = "Дата рождения";
            dataGridView3.Columns["PetName"].HeaderText = "Кличка";
            dataGridView3.Columns["Client"].HeaderText = "Хозяин";
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            var vaccine = db.Vaccines.Where(r => EF.Functions.Like(r.Name, String.Format("%" + textBox4.Text + "%"))).ToList();
            dataGridView4.DataSource = vaccine;
            dataGridView4.Columns["Id"].Visible = false;
            dataGridView4.Columns["Name"].HeaderText = "Название вакцины";
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            int index = dataGridView1.SelectedRows[0].Index;
            int id = 0;
            bool converted = Int32.TryParse(dataGridView1[0, index].Value.ToString(), out id);
            if (converted == false)
                return;

            //label2.Visible = true;
            try
            {
                Appointment appointment = db.Appointments.Find(id);
                Form lookAppointmentForm = new LookAppointmentForm(appointment.Date.ToString(),
                    appointment.Client != null ? appointment.Client.Name : "",
                    appointment.Pet != null ? appointment.Pet.PetName : "",
                    appointment.Vaccine != null ? appointment.Vaccine.Name : "", 
                    appointment.Service, appointment.WhatHurt, appointment.WhatWasDone, appointment.WhatNeedToDo);
                lookAppointmentForm.ShowDialog();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Для данного приёма не указана вакцина. Укажите в записи о приёме в поле \"Вакцина\" значение \"Без вакцины\"");
            }

        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application exApp = new Excel.Application();
                exApp.Workbooks.Add();
                Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                int i, j;
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    wsh.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                }
                for (i = 0; i <= dataGridView1.RowCount - 1; i++)
                {
                    for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                    {
                        //wsh.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                        if (dataGridView1[j, i].Value == null)
                            wsh.Cells[i + 2, j + 1] = "";
                        else
                            wsh.Cells[i + 2, j + 1] = dataGridView1[j, i].Value.ToString();
                    }
                }


                exApp.Visible = true;
                //exApp.UserControl = true;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Одно из полей \"Клиент\", \"Питомец\" или \"Вакцина\" не заполнено. Проверьте заполнение.");
            }

        }

        private void button14_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int i, j;
            for (i = 0; i <= dataGridView2.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                {
                    wsh.Cells[i + 1, j + 1] = dataGridView2[j, i].Value.ToString();
                }
            }


            exApp.Visible = true;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int i, j;
            for (j = 0; j <= dataGridView3.ColumnCount - 1; j++)
            {
                wsh.Cells[1, j + 1] = dataGridView3.Columns[j].HeaderText.ToString();
            }
            for (i = 0; i <= dataGridView3.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView3.ColumnCount - 1; j++)
                {
                    //wsh.Cells[1, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                    if (dataGridView3[j, i].Value == null)
                        wsh.Cells[i + 2, j + 1] = "";
                    else
                        wsh.Cells[i + 2, j + 1] = dataGridView3[j, i].Value.ToString();
                }
            }


            exApp.Visible = true;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            exApp.Workbooks.Add();
            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
            int i, j;
            for (i = 0; i <= dataGridView4.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView4.ColumnCount - 1; j++)
                {
                    wsh.Cells[i + 1, j + 1] = dataGridView4[j, i].Value.ToString();
                }
            }


            exApp.Visible = true;
        }
    }
}
