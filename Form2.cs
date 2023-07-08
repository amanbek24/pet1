using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace Disertation
{
    
    public partial class Form2 : Form
    {
        
        OleDbConnection conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\AMO\\source\\repos\\Disertation\\disertationbase.mdb");
      
        public Form2()
        {
            InitializeComponent();

            tabControl1.DrawItem += new DrawItemEventHandler(tabControl1_DrawItem);
            tabControl1.BackColor = Color.LightBlue;
            
        }
        
        private void tabControl1_DrawItem(Object sender, System.Windows.Forms.DrawItemEventArgs e)
        {
            Graphics g = e.Graphics;
            Brush _textBrush;

            // Get the item from the collection.
            TabPage _tabPage = tabControl1.TabPages[e.Index];

            // Get the real bounds for the tab rectangle.
            Rectangle _tabBounds = tabControl1.GetTabRect(e.Index);

            
            if (e.State == DrawItemState.Selected)
            {

                // Draw a different background color, and don't paint a focus rectangle.
                _textBrush = new SolidBrush(Color.Black);
                g.FillRectangle(Brushes.AliceBlue, e.Bounds);

            }
            else
            {
                _textBrush = new System.Drawing.SolidBrush(e.ForeColor);
                e.DrawBackground();
            }

            // Use our own font.
            Font _tabFont = new Font("Segoe UI", 13.0f, FontStyle.Bold, GraphicsUnit.Pixel);

            // Draw string. Center the text.
            StringFormat _stringFlags = new StringFormat();
            _stringFlags.Alignment = StringAlignment.Center;
            _stringFlags.LineAlignment = StringAlignment.Center;
            g.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            // Load data into the ListBox when the form loads
            LoadData();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            // Reload data when the first radio button is checked
            LoadData();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            // Reload data when the second radio button is checked
            LoadData();
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            LoadData();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            LoadData();
        }
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            LoadData();
        }
        private void Form2_Load_1(object sender, EventArgs e)
        {
            conn.Open();
            DataTable teachingStaffTable1 = new DataTable();
            string sqlQuery28 = "SELECT Full_name, dip_rus, dip_kaz, nirm_one, nirm_two, nird_one, nird_two, nird_three FROM Teaching_staff WHERE Id_step NOT IN (8, 10);";
            using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(sqlQuery28, conn))
                
            {
                adapter1.Fill(teachingStaffTable1);
            }

            foreach (DataRow row in teachingStaffTable1.Rows)
            {
                string teacherName1 = row["Full_name"].ToString();
                string nirm1 = row["nirm_one"].ToString();
                string nirm2 = row["nirm_two"].ToString();
                int nirm1Value, nirm2Value;
                int total;

                if (int.TryParse(nirm1, out nirm1Value) && int.TryParse(nirm2, out nirm2Value))
                {
                    total = nirm1Value + nirm2Value;
                }
                else
                {
                    // Handle the case where the conversion fails
                    total = 0; // or any other default value you want
                }
                DataGridViewRow dataGridViewRow = new DataGridViewRow();
                dataGridViewRow.CreateCells(dataGridView3, teacherName1, nirm1, nirm2 , total);
                dataGridView8.Rows.Add(dataGridViewRow);
                dataGridView8.ReadOnly = true;


            }
            foreach (DataRow row in teachingStaffTable1.Rows)
            {
                string teacherName1 = row["Full_name"].ToString();
                string nird1 = row["nird_one"].ToString();
                string nird2 = row["nird_two"].ToString();
                string nird3 = row["nird_three"].ToString();

                double nird1Value, nird2Value, nird3Value;
                double total1;

                if (double.TryParse(nird1, out nird1Value) && double.TryParse(nird2, out nird2Value) && double.TryParse(nird3, out nird3Value))
                {
                    total1 = nird1Value + nird2Value + nird3Value;
                }
                else
                {
                    // Handle the case where the conversion fails
                    total1 = 0; // or any other default value you want
                }

                DataGridViewRow newDataRow = new DataGridViewRow();
                newDataRow.CreateCells(dataGridView9, teacherName1, nird1, nird2, nird3, total1);
                dataGridView9.Rows.Add(newDataRow);
            }

            dataGridView9.ReadOnly = true;

            // Assuming you have two DataGridView controls named dataGridView4 and dataGridView8

            // Calculate the sum of the data from the last column of dataGridView8
            int sum = 0;
            int sum1 = 0;
            int columnIndex = 1; // Assuming the last column index is the desired column
            int columnIndex1 = 2; // Assuming the second-to-last column index is the desired column

            foreach (DataGridViewRow row in dataGridView8.Rows)
            {
                if (row.Cells[columnIndex].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex].Value.ToString(), out cellValue))
                    {
                        sum += cellValue;
                    }
                }

                if (row.Cells[columnIndex1].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex1].Value.ToString(), out cellValue))
                    {
                        sum1 += cellValue;
                    }
                }
            }
            int sub = sum - 35;
            int sub1 = sum1 - 22;
            // Add a new row to dataGridView4 with the sums as the first and second columns' values
            dataGridView4.Rows.Add("Всего", sum.ToString(), sum1.ToString());
            dataGridView4.Rows.Add("Вакансий", sub.ToString(), sub1.ToString());
            // Assuming you have two DataGridView controls named dataGridView4 and dataGridView8

            // Calculate the sum of the data from the last column of dataGridView8
            int sum2 = 0;
            int sum3 = 0;
            int sum4 = 0;
            int columnIndex2 = 1; // Assuming the last column index is the desired column
            int columnIndex3 = 2;
            int columnIndex4 = 3;// Assuming the second-to-last column index is the desired column

            foreach (DataGridViewRow row in dataGridView9.Rows)
            {
                if (row.Cells[columnIndex2].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex2].Value.ToString(), out cellValue))
                    {
                        sum2 += cellValue;
                    }
                }

                if (row.Cells[columnIndex3].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex1].Value.ToString(), out cellValue))
                    { 
                        sum3 += cellValue;
                    }
                }
                if (row.Cells[columnIndex4].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex4].Value.ToString(), out cellValue))
                    {
                        sum4 += cellValue;
                    }
                }
            }
            int sub2 = sum2 - 3;
            int sub3 = sum3 - 6;
            int sub4 = sum4 - 13;
            // Add a new row to dataGridView4 with the sums as the first and second columns' values
            dataGridView5.Rows.Add("Всего", sum2.ToString(), sum3.ToString(), sum4.ToString());
            dataGridView5.Rows.Add("Вакансий", sub2.ToString(), sub3.ToString(), sub4.ToString());





            foreach (DataRow row in teachingStaffTable1.Rows)
            {
                string teacherName1 = row["Full_name"].ToString();
                string dip1 = row["dip_rus"].ToString();
                string dip2 = row["dip_kaz"].ToString();
                int dip1Value, dip2Value;
                int total2;

                if (int.TryParse(dip1, out dip1Value) && int.TryParse(dip2, out dip2Value))
                {
                    total2 = dip1Value + dip2Value;
                }
                else
                {
                    // Handle the case where the conversion fails
                    total2 = 0; // or any other default value you want
                }
                DataGridViewRow dataGridViewRow = new DataGridViewRow();
                dataGridViewRow.CreateCells(dataGridView3, teacherName1, dip1, dip2, total2);
                dataGridView10.Rows.Add(dataGridViewRow);
                dataGridView10.ReadOnly = true;

            }
            // Calculate the sum of the data from the last column of dataGridView8
            int sum5 = 0;
            int sum6 = 0;
            int columnIndex5 = 1; // Assuming the last column index is the desired column
            int columnIndex6 = 2; // Assuming the second-to-last column index is the desired column

            foreach (DataGridViewRow row in dataGridView10.Rows)
            {
                if (row.Cells[columnIndex5].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex5].Value.ToString(), out cellValue))
                    {
                        sum5 += cellValue;
                    }
                }

                if (row.Cells[columnIndex6].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex6].Value.ToString(), out cellValue))
                    {
                        sum6 += cellValue;
                    }
                }
            }
            int sub5 = sum5 - 32;
            int sub6 = sum6 - 154;
            // Add a new row to dataGridView4 with the sums as the first and second columns' values
            dataGridView6.Rows.Add("Всего", sum5.ToString(), sum6.ToString());
            dataGridView6.Rows.Add("Вакансий", sub5.ToString(), sub6.ToString());
            // Set up data binding for dataGridView8




            // Load teaching staff data into a DataTable
            DataTable teachingStaffTable = new DataTable();
            string sqlQuery = "SELECT Full_name, rates, plan ,sem_one,sem_two, sem_one_aud, sem_two_aud,aud_total,total_k FROM Teaching_staff;";
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery, conn))
            {
                adapter.Fill(teachingStaffTable);
            }

            // Populate dataGridView3 with Full_name and total_k columns
            foreach (DataRow row in teachingStaffTable.Rows)
            {
                string teacherName = row["Full_name"].ToString();
                string rates = row["rates"].ToString();
                string plan = row["plan"].ToString();
                string total_k = row["total_k"].ToString();
                DataGridViewRow dataGridViewRow = new DataGridViewRow();
                dataGridViewRow.CreateCells(dataGridView3, teacherName, total_k, rates,plan);
                dataGridView3.Rows.Add(dataGridViewRow);
            }

            // Populate dataGridView7 with Full_name and need columns
            foreach (DataRow row in teachingStaffTable.Rows)
            {
                string teacherName = row["Full_name"].ToString();
                string rates = row["rates"].ToString();
                string plan = row["plan"].ToString();
                string total_k = row["total_k"].ToString();
                string sem_one = row["sem_one"].ToString();
                string sem_two = row["sem_two"].ToString();
                string sem_one_aud = row["sem_one_aud"].ToString();
                string sem_two_aud = row["sem_two_aud"].ToString();
                string aud_total = row["aud_total"].ToString();
                DataGridViewRow dataGridViewRow = new DataGridViewRow();
                dataGridViewRow.CreateCells(dataGridView7, teacherName, rates, sem_one, sem_one_aud, sem_two, sem_two_aud, total_k, aud_total,null, plan);
                dataGridView7.Rows.Add(dataGridViewRow);
            }
            string sqlQuery140 = "SELECT Full_name FROM Teaching_staff WHERE Id_TS = 29;";

            // Create the OleDbCommand object and set the SQL query and connection
            OleDbCommand cmd21 = new OleDbCommand(sqlQuery140, conn);

            // Create the OleDbDataReader object and execute the SQL query
            OleDbDataReader reader40 = cmd21.ExecuteReader();

            // Clear the list box and combo box
            
            comboBox14.Items.Clear();
            comboBox9.Items.Clear();


            // Loop through the results and add them to the list box and combo box
            while (reader40.Read())
            {
                
                string name = reader40.GetString(0);
                comboBox9.Items.Add(name);
                string specificName9 = "Дюсекеев К.А.";

                // Check if the specific name exists in comboBox5
                if (comboBox9.Items.Contains(specificName9))
                {
                    // Set the selected item to the specific name
                    comboBox9.SelectedItem = specificName9;
                }
                comboBox14.Items.Add(name);
                string specificName14 = "Дюсекеев К.А.";

                // Check if the specific name exists in comboBox5
                if (comboBox14.Items.Contains(specificName14))
                {
                    // Set the selected item to the specific name
                    comboBox14.SelectedItem = specificName14;
                }

            }


            reader40.Close();
            string sqlQuery40 = "SELECT Full_name FROM Teaching_staff;";

            // Create the OleDbCommand object and set the SQL query and connection
            OleDbCommand cmd20 = new OleDbCommand(sqlQuery40, conn);

            // Create the OleDbDataReader object and execute the SQL query
            OleDbDataReader reader41 = cmd20.ExecuteReader();

            // Clear the list box and combo box

            comboBox5.Items.Clear();
            comboBox6.Items.Clear();
            comboBox7.Items.Clear();
            comboBox8.Items.Clear();


            // Loop through the results and add them to the list box and combo box
            while (reader41.Read())
            {
                
                string name1 = reader41.GetString(0);
                
                comboBox5.Items.Add(name1);
                string specificName1 = "Ұзаққызы Н.";

                // Check if the specific name exists in comboBox5
                if (comboBox5.Items.Contains(specificName1))
                {
                    // Set the selected item to the specific name
                    comboBox5.SelectedItem = specificName1;
                }
                comboBox6.Items.Add(name1);
                string specificName = "Сыздыкова А. М.";

                // Check if the specific name exists in comboBox5
                if (comboBox6.Items.Contains(specificName))
                {
                    // Set the selected item to the specific name
                    comboBox6.SelectedItem = specificName;
                }
                comboBox7.Items.Add(name1);
                string specificName2 = "Нуржанова А. Б.";

                // Check if the specific name exists in comboBox5
                if (comboBox7.Items.Contains(specificName2))
                {
                    // Set the selected item to the specific name
                    comboBox7.SelectedItem = specificName2;
                }
                comboBox8.Items.Add(name1);
                string specificName3 = "Баенова Г.М.";

                // Check if the specific name exists in comboBox5
                if (comboBox8.Items.Contains(specificName3))
                {
                    // Set the selected item to the specific name
                    comboBox8.SelectedItem = specificName3;
                }
            }
           

            reader41.Close();
            string sqlQuery141 = "SELECT Full_name FROM Teaching_staff;";

            // Create the OleDbCommand object and set the SQL query and connection
            OleDbCommand cmd22 = new OleDbCommand(sqlQuery141, conn);

            // Create the OleDbDataReader object and execute the SQL query
            OleDbDataReader reader42 = cmd22.ExecuteReader();

            // Clear the list box and combo box

            comboBox10.Items.Clear();
            comboBox11.Items.Clear();
            comboBox15.Items.Clear();
            comboBox16.Items.Clear();
            comboBox12.Items.Clear();
            comboBox13.Items.Clear();
            comboBox17.Items.Clear();
            comboBox23.Items.Clear();


            // Loop through the results and add them to the list box and combo box
            while (reader42.Read())
            {
                string name2 = reader42.GetString(0);
                comboBox10.Items.Add(name2);
                string specificName10 = "Дуйсенова Г.А.";

                // Check if the specific name exists in comboBox5
                if (comboBox10.Items.Contains(specificName10))
                {
                    // Set the selected item to the specific name
                    comboBox10.SelectedItem = specificName10;
                }
                comboBox11.Items.Add(name2);
                string specificName11 = "Жартыбаева М.Г.";

                // Check if the specific name exists in comboBox5
                if (comboBox11.Items.Contains(specificName11))
                {
                    // Set the selected item to the specific name
                    comboBox11.SelectedItem = specificName11;
                }
                comboBox15.Items.Add(name2);
                string specificName15 = "Жартыбаева М.Г.";

                // Check if the specific name exists in comboBox5
                if (comboBox15.Items.Contains(specificName15))
                {
                    // Set the selected item to the specific name
                    comboBox15.SelectedItem = specificName15;
                }
                comboBox16.Items.Add(name2);
                string specificName16 = "Мирғалиқызы Т.";

                // Check if the specific name exists in comboBox5
                if (comboBox16.Items.Contains(specificName16))
                {
                    // Set the selected item to the specific name
                    comboBox16.SelectedItem = specificName16;
                }
                comboBox12.Items.Add(name2);
                string specificName12 = "Есенгалиева Ж.С.";

                // Check if the specific name exists in comboBox5
                if (comboBox12.Items.Contains(specificName12))
                {
                    // Set the selected item to the specific name
                    comboBox12.SelectedItem = specificName12;
                }
                comboBox13.Items.Add(name2);
                string specificName13 = "Толегенова Г.Б.";

                // Check if the specific name exists in comboBox5
                if (comboBox13.Items.Contains(specificName13))
                {
                    // Set the selected item to the specific name
                    comboBox13.SelectedItem = specificName13;
                }
                comboBox17.Items.Add(name2);
                string specificName17 = "Ұзаққызы Н.";

                // Check if the specific name exists in comboBox5
                if (comboBox17.Items.Contains(specificName17))
                {
                    // Set the selected item to the specific name
                    comboBox17.SelectedItem = specificName17;
                }
                comboBox23.Items.Add(name2);
                string specificName23 = "Сыздыкова А. М.";

                // Check if the specific name exists in comboBox5
                if (comboBox23.Items.Contains(specificName23))
                {
                    // Set the selected item to the specific name
                    comboBox23.SelectedItem = specificName23;
                }
            }

            
            reader42.Close();
          

            


            conn.Close();

           

        }

        private void radioButton22_CheckedChanged(object sender, EventArgs e)
        {
            conn.Open();
            
            dataGridView11.ReadOnly = true;
            DataTable teachingStaffTable = new DataTable();
            string sqlQuery888 = "SELECT appoint FROM Appointment;";
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery888, conn))
            {
                adapter.Fill(teachingStaffTable);
            }

            // Clear existing columns and rows
            dataGridView11.Columns.Clear();
            dataGridView11.Rows.Clear();

            // Create columns
            foreach (DataColumn column in teachingStaffTable.Columns)
            {
                DataGridViewTextBoxColumn gridColumn = new DataGridViewTextBoxColumn();
                gridColumn.HeaderText = column.ColumnName;
                gridColumn.Name = column.ColumnName;
                gridColumn.Visible = false; // Initially hide the column
                gridColumn.Width = dataGridView11.Width / teachingStaffTable.Columns.Count; // Set column width
                dataGridView11.Columns.Add(gridColumn);
            }

            // Populate the data
            foreach (DataRow row in teachingStaffTable.Rows)
            {
                DataGridViewRow dataGridViewRow = new DataGridViewRow();

                foreach (DataColumn column in teachingStaffTable.Columns)
                {
                    string value = row[column.ColumnName].ToString();

                    DataGridViewCell cell = new DataGridViewTextBoxCell();
                    cell.Value = value;
                    dataGridViewRow.Cells.Add(cell);

                    if (!string.IsNullOrEmpty(value))
                    {
                        dataGridView11.Columns[column.ColumnName].Visible = true; // Show the column if it has a value
                    }
                }

                dataGridView11.Rows.Add(dataGridViewRow);
            }

            conn.Close();
        }

        private void radioButton23_CheckedChanged(object sender, EventArgs e)
        {
            conn.Open();

            dataGridView11.ReadOnly = true;
            DataTable teachingStaffTable = new DataTable();
            string sqlQuery888 = "SELECT stepen FROM Degree;";
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery888, conn))
            {
                adapter.Fill(teachingStaffTable);
            }

            // Clear existing columns and rows
            dataGridView11.Columns.Clear();
            dataGridView11.Rows.Clear();

            // Create columns
            foreach (DataColumn column in teachingStaffTable.Columns)
            {
                DataGridViewTextBoxColumn gridColumn = new DataGridViewTextBoxColumn();
                gridColumn.HeaderText = column.ColumnName;
                gridColumn.Name = column.ColumnName;
                gridColumn.Visible = false; // Initially hide the column
                gridColumn.Width = dataGridView11.Width / teachingStaffTable.Columns.Count; // Set column width
                dataGridView11.Columns.Add(gridColumn);
            }

            // Populate the data
            foreach (DataRow row in teachingStaffTable.Rows)
            {
                DataGridViewRow dataGridViewRow = new DataGridViewRow();

                foreach (DataColumn column in teachingStaffTable.Columns)
                {
                    string value = row[column.ColumnName].ToString();

                    DataGridViewCell cell = new DataGridViewTextBoxCell();
                    cell.Value = value;
                    dataGridViewRow.Cells.Add(cell);

                    if (!string.IsNullOrEmpty(value))
                    {
                        dataGridView11.Columns[column.ColumnName].Visible = true; // Show the column if it has a value
                    }
                }

                dataGridView11.Rows.Add(dataGridViewRow);
            }

            conn.Close();
        }

        private void radioButton26_CheckedChanged(object sender, EventArgs e)
        {
            conn.Open();

            dataGridView11.ReadOnly = true;
            DataTable teachingStaffTable = new DataTable();
            string sqlQuery888 = "SELECT name_disciplines FROM Disciplines;";
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery888, conn))
            {
                adapter.Fill(teachingStaffTable);
            }

            // Clear existing columns and rows
            dataGridView11.Columns.Clear();
            dataGridView11.Rows.Clear();

            // Create columns
            foreach (DataColumn column in teachingStaffTable.Columns)
            {
                DataGridViewTextBoxColumn gridColumn = new DataGridViewTextBoxColumn();
                gridColumn.HeaderText = column.ColumnName;
                gridColumn.Name = column.ColumnName;
                gridColumn.Visible = false; // Initially hide the column
                gridColumn.Width = dataGridView11.Width / teachingStaffTable.Columns.Count; // Set column width
                dataGridView11.Columns.Add(gridColumn);
            }

            // Populate the data
            foreach (DataRow row in teachingStaffTable.Rows)
            {
                DataGridViewRow dataGridViewRow = new DataGridViewRow();

                foreach (DataColumn column in teachingStaffTable.Columns)
                {
                    string value = row[column.ColumnName].ToString();

                    DataGridViewCell cell = new DataGridViewTextBoxCell();
                    cell.Value = value;
                    dataGridViewRow.Cells.Add(cell);

                    if (!string.IsNullOrEmpty(value))
                    {
                        dataGridView11.Columns[column.ColumnName].Visible = true; // Show the column if it has a value
                    }
                }

                dataGridView11.Rows.Add(dataGridViewRow);
            }

            conn.Close();
        }


        private void radioButton27_CheckedChanged(object sender, EventArgs e)
        {
            conn.Open();

            dataGridView11.ReadOnly = true;
            DataTable teachingStaffTable = new DataTable();
            string sqlQuery888 =  @"SELECT ts.Full_name, a.appoint, d.stepen
                FROM (Teaching_staff ts
                INNER JOIN Appointment a ON ts.Id_App = a.Id_App)
                INNER JOIN Degree d ON ts.Id_step = d.Id_step";

            using (OleDbDataAdapter adapter = new OleDbDataAdapter(sqlQuery888, conn))
            {
                adapter.Fill(teachingStaffTable);
            }

            // Clear existing columns and rows
            dataGridView11.Columns.Clear();
            dataGridView11.Rows.Clear();

            // Create columns
            foreach (DataColumn column in teachingStaffTable.Columns)
            {
                DataGridViewTextBoxColumn gridColumn = new DataGridViewTextBoxColumn();
                gridColumn.HeaderText = column.ColumnName;
                gridColumn.Name = column.ColumnName;
                gridColumn.Width = dataGridView11.Width / teachingStaffTable.Columns.Count; // Set column width
                dataGridView11.Columns.Add(gridColumn);
            }

            // Populate the data
            foreach (DataRow row in teachingStaffTable.Rows)
            {
                object[] values = row.ItemArray;

                // Add a new row with the values
                dataGridView11.Rows.Add(values);
            }

            conn.Close();
        }





        private void LoadData()
        {
            

            int semester = 1;
           
            // Determine which radio button is checked and set the semester accordingly
            if (radioButton2.Checked)
            {
                semester = 2;
            }
            
            try
            {
                // Open the connection to the database
                conn.Open();

                // Create the SQL query
                string sqlQuery = "SELECT Disciplines.name_disciplines, Load.Id_Disc " +
                   "FROM Load " +
                   "INNER JOIN Disciplines ON Load.Id_Disc = Disciplines.Id_Disc " +
                   $"WHERE Load.semester = {semester}";
                if (radioButton4.Checked)
                {
                    sqlQuery += " AND Load.Bmd = 'B'";
                }
                else if (radioButton3.Checked)
                {
                    sqlQuery += " AND Load.Bmd = 'M'";
                }
                else if (radioButton5.Checked)
                {
                    sqlQuery += " AND Load.Bmd = 'D'";
                }
                sqlQuery += ";";
                // Create the OleDbCommand object and set the SQL query and connection
                OleDbCommand cmd = new OleDbCommand(sqlQuery, conn);

                // Create the OleDbDataReader object and execute the SQL query
                OleDbDataReader reader = cmd.ExecuteReader();

                // Clear the list box
                listBox1.Items.Clear();

                // Loop through the results and add them to the list box
                while (reader.Read())
                {
                    listBox1.Items.Add(reader.GetString(0));
                }

                // Close the data reader and the database connection
                reader.Close();

            }
            catch (Exception ex)
            {
                // Display any errors that occurred
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            // Clear the combo box
            comboBox1.Items.Clear();

            // Loop through the items of listBox1 and add them to comboBox1
            foreach (object item in listBox1.Items)
            {
                comboBox1.Items.Add(item);
            }

            // Select the first item in comboBox1
            if (comboBox1.Items.Count > 0)
            {
                comboBox1.SelectedIndex = 0;
            }
            
            string sqlQuery2 = "SELECT Full_name FROM Teaching_staff WHERE Id_TS = 65;";

            // Create the OleDbCommand object and set the SQL query and connection
            OleDbCommand cmd2 = new OleDbCommand(sqlQuery2, conn);

            // Create the OleDbDataReader object and execute the SQL query
            OleDbDataReader reader2 = cmd2.ExecuteReader();

            // Clear the list box and combo box
            listBox2.Items.Clear();
            comboBox2.Items.Clear();
            
            
            // Loop through the results and add them to the list box and combo box
            while (reader2.Read())
            {
                string name = reader2.GetString(0);
                listBox2.Items.Add(name);
                comboBox2.Items.Add(name);
                
            }
            

            reader2.Close();


           
            conn.Close();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Clear the DataGridView if the button is pressed again
            

            // Get the selected items from the ComboBoxes
            string disc = comboBox1.SelectedItem.ToString();
            string teacher = comboBox2.SelectedItem.ToString();

            // Find the index of the next empty row in the DataGridView
            int nextRowIndex = dataGridView1.Rows.Count;

            // Add a new row to the DataGridView with the selected items
            dataGridView1.Rows.Add(disc, teacher, "", "", "");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1 != null && dataGridView3 != null)
            {
                // Loop through each row in dataGridView3
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName = row.Cells["teacher1"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow = dataGridView1.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[1].Value?.ToString() == teacherName);

                        if (matchingRow != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total = Convert.ToDouble(matchingRow.Cells["lecture"].Value ?? 0)
                                + Convert.ToDouble(matchingRow.Cells["practice"].Value ?? 0)
                                + Convert.ToDouble(matchingRow.Cells["laboratory"].Value ?? 0);
                             // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            double k_now;
                            if (double.TryParse(row.Cells["k_now"].Value?.ToString(), out k_now))
                            {
                                row.Cells["k_now"].Value = k_now + total;
                            }
                            else
                            {
                                // handle the case when the value is not a valid double
                                row.Cells["k_now"].Value = total; // set the value to the total
                            }
                        }
                    }
                }
                foreach (DataGridViewRow row1 in dataGridView7.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName1 = row1.Cells["teacher2"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName1))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow1 = dataGridView1.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[1].Value?.ToString() == teacherName1);

                        if (matchingRow1 != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total1 = Convert.ToDouble(matchingRow1.Cells["lecture"].Value ?? 0)
                                + Convert.ToDouble(matchingRow1.Cells["practice"].Value ?? 0)
                                + Convert.ToDouble(matchingRow1.Cells["laboratory"].Value ?? 0);
                            total1 /= 15; // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            row1.Cells["total"].Value = Convert.ToDouble(row1.Cells["total"].Value ?? 0) + total1;
                            double lecture = Convert.ToDouble(matchingRow1.Cells["lecture"].Value ?? 0);
                            double lecTotal = Convert.ToDouble(row1.Cells["lec_total"].Value ?? 0);
                            row1.Cells["lec_total"].Value = lecTotal + lecture / 15;



                            if (radioButton1.Checked)
                            {
                                // Add the calculated total to the existing value of sem_1_aud for this row
                                row1.Cells["sem_1_aud"].Value = Convert.ToDouble(row1.Cells["sem_1_aud"].Value ?? 0) + total1;

                                // Add the calculated total to the existing value of sem_1_total for this row
                                row1.Cells["sem_1_total"].Value = Convert.ToDouble(row1.Cells["sem_1_total"].Value ?? 0) + total1;

                                row1.Cells["total_aud"].Value = Convert.ToDouble(row1.Cells["total_aud"].Value ?? 0) + total1;
                            }
                            else if (radioButton2.Checked)
                            {
                                // Add the calculated total to the existing value of sem_2_aud for this row
                                row1.Cells["sem_2_aud"].Value = Convert.ToDouble(row1.Cells["sem_2_aud"].Value ?? 0) + total1;

                                // Add the calculated total to the existing value of sem_2_total for this row
                                row1.Cells["sem_2_total"].Value = Convert.ToDouble(row1.Cells["sem_2_total"].Value ?? 0) + total1;

                                row1.Cells["total_aud"].Value = Convert.ToDouble(row1.Cells["total_aud"].Value ?? 0) + total1;
                            }
                        }

                    }

                }
            }



            
                







            
             
                    // Clear the contents of dataGridView1
                    dataGridView1.Rows.Clear();
            
        }
    


        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            e.CellStyle.BackColor = Color.AntiqueWhite;

            // Check if the current cell is a column header
            if (e.RowIndex == -1 && e.ColumnIndex > -1)
            {
                // Get the column header cell
                DataGridViewColumnHeaderCell headerCell = dataGridView1.Columns[e.ColumnIndex].HeaderCell;

                // Set the background color of the header cell
                headerCell.Style.BackColor = Color.AntiqueWhite;
            }
        }







































        OleDbConnection conne = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\AMO\\source\\repos\\Disertation\\disertationbase.mdb");

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }
        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }
        private void radioButton16_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton15_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton14_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }

        private void radioButton17_CheckedChanged(object sender, EventArgs e)
        {
            LoadData1();
        }
        private void LoadData1()
        {
            int semester = 1;

            // Determine which radio button is checked and set the semester accordingly
            if (radioButton9.Checked)
            {
                semester = 2;
            }

            try
            {
                // Open the connection to the database
                conne.Open();

                // Create the SQL query
                string sqlQuery3 = "SELECT Disciplines.name_disciplines, (Load.groups + Load.subgroups) AS total, (Load.planl + Load.plans + Load.planlab) AS credit, Load.plan_of_lecture , Load.planl, Load.plan_of_practise_exercise , Load.plans, Load.plan_of_lab, Load.planlab " +
                    "FROM Load " +
                    "INNER JOIN Disciplines ON Load.Id_Disc = Disciplines.Id_Disc " +
                    $"WHERE Load.semester = {semester}";

                if (radioButton13.Checked)
                {
                    sqlQuery3 += " AND Load.Bmd = 'B'";
                }
                else if (radioButton12.Checked)
                {
                    sqlQuery3 += " AND Load.Bmd = 'M'";
                }
                else if (radioButton11.Checked)
                {
                    sqlQuery3 += " AND Load.Bmd = 'D'";
                }

                if (radioButton8.Checked)
                {
                    sqlQuery3 += " AND Load.lang = 'K'";
                }
                else if (radioButton7.Checked)
                {
                    sqlQuery3 += " AND Load.lang = 'R'";
                }
                else if (radioButton6.Checked)
                {
                    sqlQuery3 += " AND Load.lang = 'E'";
                }

                int course = 0;
                if (radioButton16.Checked)
                {
                    course = 1;
                }
                else if (radioButton15.Checked)
                {
                    course = 2;
                }
                else if (radioButton14.Checked)
                {
                    course = 3;
                }
                else if (radioButton17.Checked)
                {
                    course = 4;
                }

                if (course > 0)
                {
                    sqlQuery3 += $" AND Load.course = {course}";
                }

                sqlQuery3 += ";";

                OleDbCommand cmd1 = new OleDbCommand(sqlQuery3, conne);

                // Create the OleDbDataReader object and execute the SQL query
                using (OleDbDataReader reader1 = cmd1.ExecuteReader())
                {
                    // Clear the data grid view
                    dataGridView15.Rows.Clear();
                    dataGridView16.Rows.Clear();

                    // Loop through the results and add them to the data grid view
                    while (reader1.Read())
                    {
                        string disciplineName = reader1.GetString(0);
                        double total = reader1.GetDouble(1);
                        double credit = reader1.GetDouble(2);
                        double plan_of_lecture = reader1.GetDouble(3);
                        double planl = reader1.GetDouble(4);
                        double plan_of_practise_exercise = reader1.GetDouble(5);
                        double plans = reader1.GetDouble(6);
                        double plan_of_lab = reader1.GetDouble(7);
                        double planlab = reader1.GetDouble(8);

                        dataGridView15.Rows.Add(disciplineName, credit, total, plan_of_lecture, planl, plan_of_practise_exercise, plans, plan_of_lab, planlab, null);
                    }
                }
            }
            catch (Exception ex)
            {
                // Display any errors that occurred
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                // Close the database connection
                conne.Close();
            }

            // Clear the combo box
            comboBox4.Items.Clear();

            // Loop through the items of dataGridView15 and add them to comboBox4
            foreach (DataGridViewRow row in dataGridView15.Rows)
            {
                if (!row.IsNewRow)
                {
                    comboBox4.Items.Add(row.Cells[0].Value);
                }
            }

            // Select the first item in comboBox4
            if (comboBox4.Items.Count > 0)
            {
                comboBox4.SelectedIndex = 0;
            }

            

            

            // Close the database connection
            
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Get the selected discipline name from the comboBox
            string disciplineName = comboBox4.SelectedItem.ToString();

            // Clear the rows in dataGridView16
            dataGridView16.Rows.Clear();

            // Loop through the results in dataGridView15 to find the relevant discipline
            foreach (DataGridViewRow row in dataGridView15.Rows)
            {
                if (!row.IsNewRow && row.Cells[0].Value.ToString() == disciplineName)
                {
                    // Retrieve the corresponding plan_of_lecture, planl, plan_of_practise_exercise, plans, plan_of_lab, and planlab values
                    double plan_of_lecture = Convert.ToDouble(row.Cells[3].Value);
                    double planl = Convert.ToDouble(row.Cells[4].Value);  // Update the cell index
                    double plan_of_practise_exercise = Convert.ToDouble(row.Cells[5].Value);
                    double plans = Convert.ToDouble(row.Cells[6].Value);  // Update the cell index
                    double plan_of_lab = Convert.ToDouble(row.Cells[7].Value);
                    double planlab = Convert.ToDouble(row.Cells[8].Value); // Update the cell index

                    // Add the values to dataGridView16
                    dataGridView16.Rows.Add(plan_of_lecture, planl, plan_of_practise_exercise, plans, plan_of_lab, planlab, null);
                }
            }
            // Get the selected discipline name from the comboBox
            

            // Clear the rows in dataGridView17
            dataGridView17.Rows.Clear();

            try
            {
                // Open the connection to the database
                conne.Open();

                // Create the SQL query to retrieve the data based on the selected discipline name
                string sqlQuery = "SELECT potoch.FIO, potoch.potoc, potoch.colg, potoch.colpg " +
                                  "FROM potoch " +
                                  "INNER JOIN Disciplines ON potoch.Id_Disc = Disciplines.Id_Disc " +
                                  "WHERE Disciplines.name_disciplines = @DisciplineName";

                OleDbCommand cmd = new OleDbCommand(sqlQuery, conne);
                cmd.Parameters.AddWithValue("@DisciplineName", disciplineName);

                // Create the OleDbDataReader object and execute the SQL query
                using (OleDbDataReader reader = cmd.ExecuteReader())
                {
                    // Loop through the results and add them to dataGridView17
                    while (reader.Read())
                    {
                        string fio = reader.GetString(0);
                        double potoc = reader.GetDouble(1);
                        double colg = reader.GetDouble(2);
                        double colpg = reader.GetDouble(3);


                        dataGridView17.Rows.Add(fio, potoc, colg, colpg);
                    }
                }
            }
            catch (Exception ex)
            {
                // Display any errors that occurred
                MessageBox.Show("An error occurred: " + ex.Message);
            }
            finally
            {
                // Close the database connection
                conne.Close();
            }
            dataGridView18.Rows.Clear();
            int sum5 = 0;
            int sum6 = 0;
            int sum7 = 0;
            int columnIndex5 = 1; // Assuming the last column index is the desired column
            int columnIndex6 = 2;
            int columnIndex7 = 3;// Assuming the second-to-last column index is the desired column

            foreach (DataGridViewRow row in dataGridView17.Rows)
            {
                if (row.Cells[columnIndex5].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex5].Value.ToString(), out cellValue))
                    {
                        sum5 += cellValue;
                    }
                }

                if (row.Cells[columnIndex6].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex6].Value.ToString(), out cellValue))
                    {
                        sum6 += cellValue;
                    }
                }
                if (row.Cells[columnIndex7].Value != null)
                {
                    int cellValue;
                    if (int.TryParse(row.Cells[columnIndex7].Value.ToString(), out cellValue))
                    {
                        sum7 += cellValue;
                    }
                }
            }
            int sub5 = sum5 - sum5;
            int sub6 = sum6 - sum6;
            int sub7 = sum7 - sum7;
            // Add a new row to dataGridView4 with the sums as the first and second columns' values
            dataGridView18.Rows.Add("Всего", sum5.ToString(), sum6.ToString(), sum7.ToString());
            dataGridView18.Rows.Add("Не распределенных", sub5.ToString(), sub6.ToString(), sub7.ToString());
        }




       
        private void button3_Click(object sender, EventArgs e)
        {
            // Set the DataGridView's ReadOnly property to false
            dataGridView17.ReadOnly = false;

            // Enable the AllowUserToAddRows and AllowUserToDeleteRows properties if necessary
            dataGridView17.AllowUserToAddRows = true;
            dataGridView17.AllowUserToDeleteRows = true;

            // Enable editing mode for the DataGridView
            dataGridView17.BeginEdit(true);

        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView17 != null && dataGridView3 != null)
            {
                // Loop through each row in dataGridView3
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName = row.Cells["teacher1"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName))
                    {
                        // Find the corresponding rows in dataGridView17 based on the teacher name
                        var selectedRows = dataGridView17.Rows
                            .Cast<DataGridViewRow>()
                            .Where(r => r.Cells[0].Value?.ToString() == teacherName);

                        // Calculate the total value for each selected row
                        foreach (DataGridViewRow selectedRow in selectedRows)
                        {
                            // Retrieve the values from the selected row
                            string fio = selectedRow.Cells[0].Value.ToString();
                            double potoc = Convert.ToDouble(selectedRow.Cells[1].Value);
                            double colg = Convert.ToDouble(selectedRow.Cells[2].Value);
                            double colpg = Convert.ToDouble(selectedRow.Cells[3].Value);

                            // Retrieve the corresponding plan_of_lecture, plan_of_practise_exercise, and plan_of_lab values
                            double plan_of_lecture = Convert.ToDouble(dataGridView16.Rows[selectedRow.Index].Cells[0].Value);
                            double plan_of_practise_exercise = Convert.ToDouble(dataGridView16.Rows[selectedRow.Index].Cells[2].Value);
                            double plan_of_lab = Convert.ToDouble(dataGridView16.Rows[selectedRow.Index].Cells[4].Value);

                            // Calculate the total value
                            double total3 = (potoc * plan_of_lecture) + (colg * plan_of_practise_exercise) + (colpg * plan_of_lab);

                            // Add the calculated total to the existing value of k_now for this row
                            double k_now;
                            if (double.TryParse(row.Cells["k_now"].Value?.ToString(), out k_now))
                            {
                                row.Cells["k_now"].Value = k_now + total3;
                            }
                            else
                            {
                                // handle the case when the value is not a valid double
                                row.Cells["k_now"].Value = total3; // set the value to the total
                            }
                        }
                    }
                }

                foreach (DataGridViewRow row1 in dataGridView7.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName1 = row1.Cells["teacher2"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName1))
                    {
                        // Find the corresponding rows in dataGridView17 based on the teacher name
                        var selectedRows = dataGridView17.Rows
                            .Cast<DataGridViewRow>()
                            .Where(r => r.Cells[0].Value?.ToString() == teacherName1);

                        // Calculate the total value for each selected row
                        foreach (DataGridViewRow selectedRow in selectedRows)
                        {
                            // Retrieve the values from the selected row
                            string fio = selectedRow.Cells[0].Value.ToString();
                            double potoc = Convert.ToDouble(selectedRow.Cells[1].Value);
                            double colg = Convert.ToDouble(selectedRow.Cells[2].Value);
                            double colpg = Convert.ToDouble(selectedRow.Cells[3].Value);

                            // Retrieve the corresponding plan_of_lecture, plan_of_practise_exercise, and plan_of_lab values
                            double plan_of_lecture = Convert.ToDouble(dataGridView16.Rows[selectedRow.Index].Cells[0].Value);
                            double plan_of_practise_exercise = Convert.ToDouble(dataGridView16.Rows[selectedRow.Index].Cells[2].Value);
                            double plan_of_lab = Convert.ToDouble(dataGridView16.Rows[selectedRow.Index].Cells[4].Value);

                            // Calculate the total value
                            double total4 = (potoc * plan_of_lecture) + (colg * plan_of_practise_exercise) + (colpg * plan_of_lab);

                            // Add the calculated total to the existing value of total for this row
                            row1.Cells["total"].Value = Convert.ToDouble(row1.Cells["total"].Value ?? 0) + total4;

                            // Add the calculated lecture value to the existing value of lec_total for this row
                            double lecTotal = Convert.ToDouble(row1.Cells["lec_total"].Value ?? 0);
                            row1.Cells["lec_total"].Value = lecTotal + (potoc * plan_of_lecture);

                            if (radioButton10.Checked)
                            {
                                // Add the calculated total to the existing value of sem_1_aud for this row
                                row1.Cells["sem_1_aud"].Value = Convert.ToDouble(row1.Cells["sem_1_aud"].Value ?? 0) + total4;

                                // Add the calculated total to the existing value of sem_1_total for this row
                                row1.Cells["sem_1_total"].Value = Convert.ToDouble(row1.Cells["sem_1_total"].Value ?? 0) + total4;

                                row1.Cells["total_aud"].Value = Convert.ToDouble(row1.Cells["total_aud"].Value ?? 0) + total4;
                            }
                            else if (radioButton9.Checked)
                            {
                                // Add the calculated total to the existing value of sem_2_aud for this row
                                row1.Cells["sem_2_aud"].Value = Convert.ToDouble(row1.Cells["sem_2_aud"].Value ?? 0) + total4;

                                // Add the calculated total to the existing value of sem_2_total for this row
                                row1.Cells["sem_2_total"].Value = Convert.ToDouble(row1.Cells["sem_2_total"].Value ?? 0) + total4;

                                row1.Cells["total_aud"].Value = Convert.ToDouble(row1.Cells["total_aud"].Value ?? 0) + total4;
                            }
                        }
                    }
                }
            }
        }




        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }
       

        private void dataGridView2_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        
       

        private void dataGridView8_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
         

        private void tabControl1_Selecting_1(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage == tabPage9)
            {
                // If the user clicks on tabPage9, close the form to exit the application
                this.Close();
            }
        }

        private void dataGridView7_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView3_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView8_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }
        

        private void dataGridView9_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView10_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView13_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView14_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView15_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView16_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        

        
        private void button9_Click(object sender, EventArgs e)
        {
            // Set the DataGridView's ReadOnly property to false
            dataGridView8.ReadOnly = false;

            // Enable the AllowUserToAddRows and AllowUserToDeleteRows properties if necessary
            dataGridView8.AllowUserToAddRows = true;
            dataGridView8.AllowUserToDeleteRows = true;

            // Enable editing mode for the DataGridView
            dataGridView8.BeginEdit(true);
        }

       
        private void button11_Click(object sender, EventArgs e)
        {
            // Set the DataGridView's ReadOnly property to false
            dataGridView9.ReadOnly = false;

            // Enable the AllowUserToAddRows and AllowUserToDeleteRows properties if necessary
            dataGridView9.AllowUserToAddRows = true;
            dataGridView9.AllowUserToDeleteRows = true;

            // Enable editing mode for the DataGridView
            dataGridView9.BeginEdit(true);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            // Set the DataGridView's ReadOnly property to false
            dataGridView10.ReadOnly = false;

            // Enable the AllowUserToAddRows and AllowUserToDeleteRows properties if necessary
            dataGridView10.AllowUserToAddRows = true;
            dataGridView10.AllowUserToDeleteRows = true;

            // Enable editing mode for the DataGridView
            dataGridView10.BeginEdit(true);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (dataGridView8 != null && dataGridView3 != null)
            {
                // Loop through each row in dataGridView3
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName = row.Cells["teacher1"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow7 = dataGridView8.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                        if (matchingRow7 != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total7 = Convert.ToDouble(matchingRow7.Cells["nirm_totall"].Value ?? 0);

                            total7 *= 0.65; // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            double k_now;
                            if (double.TryParse(row.Cells["k_now"].Value?.ToString(), out k_now))
                            {
                                // If the value of k_now + total7 is 0, set the value to an empty string
                                if (k_now + total7 == 0)
                                {
                                    row.Cells["k_now"].Value = "";
                                }
                                else
                                {
                                    row.Cells["k_now"].Value = k_now + total7;
                                }
                            }
                            else
                            {
                                // handle the case when the value is not a valid double
                                // If the value of total7 is 0, set the value to an empty string
                                if (total7 == 0)
                                {
                                    row.Cells["k_now"].Value = "";
                                }
                                else
                                {
                                    row.Cells["k_now"].Value = total7; // set the value to the total
                                }
                            }

                        }
                    }
                }
                foreach (DataGridViewRow row1 in dataGridView7.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName1 = row1.Cells["teacher2"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName1))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow8 = dataGridView8.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName1);

                        if (matchingRow8 != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total8 = Convert.ToDouble(matchingRow8.Cells["nirm_totall"].Value ?? 0);

                            total8 *= 0.65; // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            row1.Cells["total"].Value = Convert.ToDouble(row1.Cells["total"].Value ?? 0) + total8;

                            // Add the calculated total to the existing value of sem_1_total for this row
                            row1.Cells["sem_1_total"].Value = Convert.ToDouble(row1.Cells["sem_1_total"].Value ?? 0) + total8;

                            // Add the calculated total to the existing value of sem_2_total for this row
                            row1.Cells["sem_2_total"].Value = Convert.ToDouble(row1.Cells["sem_2_total"].Value ?? 0) + total8;
                            
                            // Set the value of the cell to null if the calculated total is 0
                            if (total8 == 0)
                            {
                                
                            }
                            
                        }

                    }

                }

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (dataGridView9 != null && dataGridView3 != null)
            {
                // Loop through each row in dataGridView3
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName = row.Cells["teacher1"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow7 = dataGridView9.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                        if (matchingRow7 != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total7 = Convert.ToDouble(matchingRow7.Cells["totality"].Value ?? 0);

                            total7 *= 0.65; // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            double k_now;
                            if (double.TryParse(row.Cells["k_now"].Value?.ToString(), out k_now))
                            {
                                // If the value of k_now + total7 is 0, set the value to an empty string
                                if (k_now + total7 == 0)
                                {
                                    row.Cells["k_now"].Value = "";
                                }
                                else
                                {
                                    row.Cells["k_now"].Value = k_now + total7;
                                }
                            }
                            else
                            {
                                // handle the case when the value is not a valid double
                                // If the value of total7 is 0, set the value to an empty string
                                if (total7 == 0)
                                {
                                    row.Cells["k_now"].Value = "";
                                }
                                else
                                {
                                    row.Cells["k_now"].Value = total7; // set the value to the total
                                }
                            }

                        }
                    }
                }
                foreach (DataGridViewRow row1 in dataGridView7.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName1 = row1.Cells["teacher2"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName1))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow8 = dataGridView9.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName1);

                        if (matchingRow8 != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total8 = Convert.ToDouble(matchingRow8.Cells["totality"].Value ?? 0);

                            total8 *= 0.65; // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            row1.Cells["total"].Value = Convert.ToDouble(row1.Cells["total"].Value ?? 0) + total8;

                            // Add the calculated total to the existing value of sem_1_total for this row
                            row1.Cells["sem_1_total"].Value = Convert.ToDouble(row1.Cells["sem_1_total"].Value ?? 0) + total8;

                            // Add the calculated total to the existing value of sem_2_total for this row
                            row1.Cells["sem_2_total"].Value = Convert.ToDouble(row1.Cells["sem_2_total"].Value ?? 0) + total8;

                            // Set the value of the cell to null if the calculated total is 0
                           
                            
                        }

                    }

                }

            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (dataGridView10 != null && dataGridView3 != null)
            {
                // Loop through each row in dataGridView3
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName = row.Cells["teacher1"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow7 = dataGridView10.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                        if (matchingRow7 != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total7 = Convert.ToDouble(matchingRow7.Cells["dip_totall"].Value ?? 0);

                            total7 *= 0.65; // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            double k_now;
                            if (double.TryParse(row.Cells["k_now"].Value?.ToString(), out k_now))
                            {
                                // If the value of k_now + total7 is 0, set the value to an empty string
                                if (k_now + total7 == 0)
                                {
                                    row.Cells["k_now"].Value = "";
                                }
                                else
                                {
                                    row.Cells["k_now"].Value = k_now + total7;
                                }
                            }
                            else
                            {
                                // handle the case when the value is not a valid double
                                // If the value of total7 is 0, set the value to an empty string
                                if (total7 == 0)
                                {
                                    row.Cells["k_now"].Value = "";
                                }
                                else
                                {
                                    row.Cells["k_now"].Value = total7; // set the value to the total
                                }
                            }

                        }
                    }
                }
                foreach (DataGridViewRow row1 in dataGridView7.Rows)
                {
                    // Get the teacher name from the current row
                    string teacherName1 = row1.Cells["teacher2"].Value?.ToString();

                    // Check if the teacher name is not null before proceeding
                    if (!string.IsNullOrEmpty(teacherName1))
                    {
                        // Find the corresponding row in dataGridView1 based on the teacher name
                        DataGridViewRow matchingRow8 = dataGridView10.Rows
                            .Cast<DataGridViewRow>()
                            .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName1);

                        if (matchingRow8 != null)
                        {
                            // Calculate the total value from the lecture, practice, and laboratory columns of the matching row
                            double total8 = Convert.ToDouble(matchingRow8.Cells["dip_totall"].Value ?? 0);

                            total8 *= 0.65; // Divide the total by 15

                            // Add the calculated total to the existing value of k_now for this row
                            row1.Cells["total"].Value = Convert.ToDouble(row1.Cells["total"].Value ?? 0) + total8;

                            // Add the calculated total to the existing value of sem_1_total for this row
                            row1.Cells["sem_1_total"].Value = Convert.ToDouble(row1.Cells["sem_1_total"].Value ?? 0) + total8;

                            // Add the calculated total to the existing value of sem_2_total for this row
                            row1.Cells["sem_2_total"].Value = Convert.ToDouble(row1.Cells["sem_2_total"].Value ?? 0) + total8;

                            // Set the value of the cell to null if the calculated total is 0
                           
                            
                        }

                    }

                }

            }

        }

        private void button19_Click(object sender, EventArgs e)
        {
            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox5.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 0.96;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 0.96;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 0.96;
                        matchingRow3.Cells["k_now"].Value = kNow + 0.96;
                    }
                    
                }
            }



            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox6.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 1.59;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 1.59;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 1.59;
                        matchingRow3.Cells["k_now"].Value = kNow + 1.59;
                    }
                    
                }
            }




            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox7.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 1.27;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 1.27;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 1.27;
                        matchingRow3.Cells["k_now"].Value = kNow + 1.27;
                    }
                    
                }
            }


            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox8.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 0.5;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 0.5;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 0.5;
                        matchingRow3.Cells["k_now"].Value = kNow + 0.5;
                    }
                }
            }
            
        }

        private void button17_Click(object sender, EventArgs e)
        {
            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox9.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 0.96;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 0.96;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 0.96;
                        matchingRow3.Cells["k_now"].Value = kNow + 0.96;
                    }
                }
            }



            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox10.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 4.02;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 4.02;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 4.02;
                        matchingRow3.Cells["k_now"].Value = kNow + 4.02;
                    }
                        
                }
            }




            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox11.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 4.02;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 4.02;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 4.02;
                        matchingRow3.Cells["k_now"].Value = kNow + 4.02;
                    }
                        
                }
            }


            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox12.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 4.02;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 4.02;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 4.02;
                        matchingRow3.Cells["k_now"].Value = kNow + 4.02;
                    }
                        
                }
            }

            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox13.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 4.02;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 4.02;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 4.02;
                        matchingRow3.Cells["k_now"].Value = kNow + 4.02;
                    }
                        
                }
            }
            
        }

        private void button15_Click(object sender, EventArgs e)
        {
            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox14.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 0.88;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 0.88;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 0.88;
                        matchingRow3.Cells["k_now"].Value = kNow + 0.88;
                    }
                        
                }
            }
            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox15.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 0.88;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 0.88;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 0.88;
                        matchingRow3.Cells["k_now"].Value = kNow + 0.88;
                    }
                    
                }
            }
            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox16.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 0.88;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 0.88;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 0.88;
                        matchingRow3.Cells["k_now"].Value = kNow + 0.88;
                    }
                }
                    
            }
            {
                // Get the selected teacher name from combobox5
                string teacherName = comboBox17.SelectedItem?.ToString();

                // Check if the teacher name is not null or empty before proceeding
                if (!string.IsNullOrEmpty(teacherName))
                {
                    // Find the corresponding row in dataGridView3 based on the teacher name
                    DataGridViewRow matchingRow3 = dataGridView3.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Find the corresponding row in dataGridView7 based on the teacher name
                    DataGridViewRow matchingRow7 = dataGridView7.Rows
                        .Cast<DataGridViewRow>()
                        .FirstOrDefault(r => r.Cells[0].Value?.ToString() == teacherName);

                    // Check if both matching rows are not null before proceeding
                    if (matchingRow3 != null && matchingRow7 != null)
                    {
                        // Calculate the total value from the columns of the matching rows
                        double total = 0;
                        double sem1Total = 0;
                        double sem2Total = 0;
                        double kNow = 0;

                        double.TryParse(matchingRow7.Cells["total"].Value?.ToString(), out total);
                        double.TryParse(matchingRow7.Cells["sem_1_total"].Value?.ToString(), out sem1Total);
                        double.TryParse(matchingRow7.Cells["sem_2_total"].Value?.ToString(), out sem2Total);
                        double.TryParse(matchingRow3.Cells["k_now"].Value?.ToString(), out kNow);

                        // Add the calculated values to the existing values of the specified columns for the matching rows
                        matchingRow7.Cells["total"].Value = total + 0.88;
                        matchingRow7.Cells["sem_1_total"].Value = sem1Total + 0.88;
                        matchingRow7.Cells["sem_2_total"].Value = sem2Total + 0.88;
                        matchingRow3.Cells["k_now"].Value = kNow + 0.88;
                    }
                        
                }
            }
            
           
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void dataGridView11_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView11_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView4_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView5_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView6_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView12_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            

        }

        private void dataGridView12_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }
        

        private void comboBox23_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            string selectedName = comboBox23.SelectedItem.ToString();

            if (selectedName == "Сыздыкова А. М.")
            {
                dataGridView12.ReadOnly = true;
                // Set the LicenseContext to suppress the LicenseException
                OfficeOpenXml.LicenseContext licenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelPackage.LicenseContext = licenseContext;

                // Specify the path to the Excel file
                string excelFilePath = @"C:\Users\AMO\source\repos\Disertation\sizdikova.xlsx";

                // Load the Excel file using EPPlus
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    // Get the first worksheet in the Excel file
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Get the total number of rows and columns in the worksheet
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // Clear the existing data from the DataGridView
                    dataGridView12.Rows.Clear();
                    dataGridView12.Columns.Clear();

                    // Loop through each column to create DataGridView columns
                    // Loop through each column to create DataGridView columns
                    for (int col = 1; col <= colCount; col++)
                    {
                        // Get the column header value from the worksheet
                        string columnHeader = worksheet.Cells[1, col].Value?.ToString();

                        // Create a DataGridViewTextBoxColumn for each column
                        DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                        column.HeaderText = columnHeader;

                        // Add the column to the DataGridView
                        dataGridView12.Columns.Add(column);

                        // Set the width of columns except the first one to 26
                        if (col > 1 && col < colCount)
                        {
                            dataGridView12.Columns[col - 1].Width = 38;
                        }
                        // Set the width of the last column to 60
                        else if (col == colCount)
                        {
                            dataGridView12.Columns[col - 1].Width = 52;
                        }

                    }




                    // Disable the ability to add new rows
                    dataGridView12.AllowUserToAddRows = false;

                    // Loop through each row to transfer the data
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Create an array to hold the row data
                        string[] rowData = new string[colCount];

                        // Loop through each column to retrieve the cell values
                        for (int col = 1; col <= colCount; col++)
                        {
                            // Get the cell value from the worksheet
                            string cellValue = worksheet.Cells[row, col].Value?.ToString();

                            // Store the cell value in the row data array
                            rowData[col - 1] = cellValue;
                        }

                        // Add the row data to the DataGridView
                        dataGridView12.Rows.Add(rowData);

                        // Hide the row if it doesn't have data in the first column
                        if (string.IsNullOrEmpty(rowData[0]))
                        {
                            dataGridView12.Rows[dataGridView12.Rows.Count - 1].Visible = false;
                        }
                    }
                }
            }
            else if (selectedName == "Сайтова Р.Б.")
            {
                dataGridView12.ReadOnly = true;
                // Set the LicenseContext to suppress the LicenseException
                OfficeOpenXml.LicenseContext licenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelPackage.LicenseContext = licenseContext;

                // Specify the path to the Excel file
                string excelFilePath = @"C:\Users\AMO\source\repos\Disertation\saytova.xlsx";

                // Load the Excel file using EPPlus
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    // Get the first worksheet in the Excel file
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Get the total number of rows and columns in the worksheet
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // Clear the existing data from the DataGridView
                    dataGridView12.Rows.Clear();
                    dataGridView12.Columns.Clear();

                    // Loop through each column to create DataGridView columns
                    // Loop through each column to create DataGridView columns
                    for (int col = 1; col <= colCount; col++)
                    {
                        // Get the column header value from the worksheet
                        string columnHeader = worksheet.Cells[1, col].Value?.ToString();

                        // Create a DataGridViewTextBoxColumn for each column
                        DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                        column.HeaderText = columnHeader;

                        // Add the column to the DataGridView
                        dataGridView12.Columns.Add(column);

                        // Set the width of columns except the first one to 26
                        if (col > 1 && col < colCount)
                        {
                            dataGridView12.Columns[col - 1].Width = 38;
                        }
                        // Set the width of the last column to 60
                        else if (col == colCount)
                        {
                            dataGridView12.Columns[col - 1].Width = 52;
                        }

                    }




                    // Disable the ability to add new rows
                    dataGridView12.AllowUserToAddRows = false;

                    // Loop through each row to transfer the data
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Create an array to hold the row data
                        string[] rowData = new string[colCount];

                        // Loop through each column to retrieve the cell values
                        for (int col = 1; col <= colCount; col++)
                        {
                            // Get the cell value from the worksheet
                            string cellValue = worksheet.Cells[row, col].Value?.ToString();

                            // Store the cell value in the row data array
                            rowData[col - 1] = cellValue;
                        }

                        // Add the row data to the DataGridView
                        dataGridView12.Rows.Add(rowData);

                        // Hide the row if it doesn't have data in the first column
                        if (string.IsNullOrEmpty(rowData[0]))
                        {
                            dataGridView12.Rows[dataGridView12.Rows.Count - 1].Visible = false;
                        }
                    }
                }
            }
            else if (selectedName == "Есенгалиева Ж.С.")
            {
                dataGridView12.ReadOnly = true;
                // Set the LicenseContext to suppress the LicenseException
                OfficeOpenXml.LicenseContext licenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                ExcelPackage.LicenseContext = licenseContext;

                // Specify the path to the Excel file
                string excelFilePath = @"C:\Users\AMO\source\repos\Disertation\esengalieva.xlsx";

                // Load the Excel file using EPPlus
                using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    // Get the first worksheet in the Excel file
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    // Get the total number of rows and columns in the worksheet
                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    // Clear the existing data from the DataGridView
                    dataGridView12.Rows.Clear();
                    dataGridView12.Columns.Clear();

                    // Loop through each column to create DataGridView columns
                    // Loop through each column to create DataGridView columns
                    for (int col = 1; col <= colCount; col++)
                    {
                        // Get the column header value from the worksheet
                        string columnHeader = worksheet.Cells[1, col].Value?.ToString();

                        // Create a DataGridViewTextBoxColumn for each column
                        DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                        column.HeaderText = columnHeader;

                        // Add the column to the DataGridView
                        dataGridView12.Columns.Add(column);

                        // Set the width of columns except the first one to 26
                        if (col > 1 && col < colCount)
                        {
                            dataGridView12.Columns[col - 1].Width = 38;
                        }
                        // Set the width of the last column to 60
                        else if (col == colCount)
                        {
                            dataGridView12.Columns[col - 1].Width = 52;
                        }

                    }




                    // Disable the ability to add new rows
                    dataGridView12.AllowUserToAddRows = false;

                    // Loop through each row to transfer the data
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Create an array to hold the row data
                        string[] rowData = new string[colCount];

                        // Loop through each column to retrieve the cell values
                        for (int col = 1; col <= colCount; col++)
                        {
                            // Get the cell value from the worksheet
                            string cellValue = worksheet.Cells[row, col].Value?.ToString();

                            // Store the cell value in the row data array
                            rowData[col - 1] = cellValue;
                        }

                        // Add the row data to the DataGridView
                        dataGridView12.Rows.Add(rowData);

                        // Hide the row if it doesn't have data in the first column
                        if (string.IsNullOrEmpty(rowData[0]))
                        {
                            dataGridView12.Rows[dataGridView12.Rows.Count - 1].Visible = false;
                        }
                    }
                }
            }
        }

        private void radioButton21_CheckedChanged(object sender, EventArgs e)
        {
            // Clearing dataGridView13
            dataGridView13.Rows.Clear();

            // Clearing dataGridView14
            dataGridView14.Rows.Clear();

            if (radioButton21.Checked)
            {
                conn.Open();
                DataTable teachingStaffTable1 = new DataTable();
                string query = "SELECT Full_name, uch_one, uch_two FROM Teaching_staff";
                using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(query, conn))
                {
                    adapter1.Fill(teachingStaffTable1);
                }

                int sum2 = 0;
                int sub2 = 0;
                int sub3 = 0;
                int sum3 = 0;

                foreach (DataRow row in teachingStaffTable1.Rows)
                {
                    string teacherName1 = row["Full_name"].ToString();
                    string uch_one = row["uch_one"].ToString();
                    string uch_two = row["uch_two"].ToString();
                    int nirm1Value, nirm2Value;
                    

                    if (int.TryParse(uch_one, out nirm1Value) && int.TryParse(uch_two, out nirm2Value))
                    {
                         
                        sum2 += nirm1Value;
                        sum3 += nirm2Value;
                        sub2 = sum2 - 1;
                        sub3 = sum3 - 3;
                    }
                    else
                    {
                        // Handle the case where the conversion fails
                        sum2 = 0; // or any other default value you want
                    }

                    DataGridViewRow dataGridViewRow = new DataGridViewRow();
                    dataGridViewRow.CreateCells(dataGridView13, teacherName1, uch_one, uch_two);
                    dataGridView13.Rows.Add(dataGridViewRow);
                    dataGridView13.ReadOnly = true;
                }

                dataGridView14.Rows.Add("Всего", sum2.ToString(), sum3.ToString());
                dataGridView14.Rows.Add("Вакансий", sub2.ToString(), sub3.ToString());

                conn.Close();
            }

        }


        private void radioButton20_CheckedChanged(object sender, EventArgs e)
        {
            // Clearing dataGridView13
            dataGridView13.Rows.Clear();

            // Clearing dataGridView14
            dataGridView14.Rows.Clear();

            if (radioButton20.Checked)
            {
                conn.Open();
                DataTable teachingStaffTable1 = new DataTable();
                string query = "SELECT Full_name, pro_one, pro_two FROM Teaching_staff";
                using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(query, conn))
                {
                    adapter1.Fill(teachingStaffTable1);
                }

                int sum2 = 0;
                int sub2 = 0;
                int sub3 = 0;
                int sum3 = 0;

                foreach (DataRow row in teachingStaffTable1.Rows)
                {
                    string teacherName1 = row["Full_name"].ToString();
                    string pro_one = row["pro_one"].ToString();
                    string pro_two = row["pro_two"].ToString();
                    int nirm1Value, nirm2Value;
                    

                    if (int.TryParse(pro_one, out nirm1Value) && int.TryParse(pro_two, out nirm2Value))
                    {
                        sum2 += nirm1Value;
                        sum3 += nirm2Value;
                        sub2 = sum2 - 2;
                        sub3 = sum3 - 5;
                        
                    }
                   

                    DataGridViewRow dataGridViewRow = new DataGridViewRow();
                    dataGridViewRow.CreateCells(dataGridView13, teacherName1, pro_one, pro_two);
                    dataGridView13.Rows.Add(dataGridViewRow);
                    dataGridView13.ReadOnly = true;
                }

                dataGridView14.Rows.Add("Всего", sum2.ToString(), sum3.ToString());
                dataGridView14.Rows.Add("Вакансий", sub2.ToString(), sub3.ToString());
                conn.Close();
            }
        }

        private void radioButton19_CheckedChanged(object sender, EventArgs e)
        {
            // Clearing dataGridView13
            dataGridView13.Rows.Clear();

            // Clearing dataGridView14
            dataGridView14.Rows.Clear();

            if (radioButton19.Checked)
            {
                conn.Open();
                DataTable teachingStaffTable1 = new DataTable();
                string query = "SELECT Full_name, pros_one, pros_two FROM Teaching_staff";
                using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(query, conn))
                {
                    adapter1.Fill(teachingStaffTable1);
                }

                int sum2 = 0;
                int sub2 = 0;
                int sub3 = 0;
                int sum3 = 0;

                foreach (DataRow row in teachingStaffTable1.Rows)
                {
                    string teacherName1 = row["Full_name"].ToString();
                    string pros_one = row["pros_one"].ToString();
                    string pros_two = row["pros_two"].ToString();
                    int nirm1Value, nirm2Value;
                    

                    if (int.TryParse(pros_one, out nirm1Value) && int.TryParse(pros_two, out nirm2Value))
                    {

                        sum2 += nirm1Value;
                        sum3 += nirm2Value;
                        sub2 = sum2 - 2;
                        sub3 = sum3 - 4;
                    }
                    else
                    {
                        // Handle the case where the conversion fails
                        sum2 = 0; // or any other default value you want
                    }

                    DataGridViewRow dataGridViewRow = new DataGridViewRow();
                    dataGridViewRow.CreateCells(dataGridView13, teacherName1, pros_one, pros_two);
                    dataGridView13.Rows.Add(dataGridViewRow);
                    dataGridView13.ReadOnly = true;
                }

                dataGridView14.Rows.Add("Всего", sum2.ToString(), sum3.ToString());
                dataGridView14.Rows.Add("Вакансий", sub2.ToString(), sub3.ToString());
                conn.Close();
            }
        }

        private void radioButton18_CheckedChanged(object sender, EventArgs e)
        {
            // Clearing dataGridView13
            dataGridView13.Rows.Clear();

            // Clearing dataGridView14
            dataGridView14.Rows.Clear();

            if (radioButton18.Checked)
            {
                conn.Open();
                DataTable teachingStaffTable1 = new DataTable();
                string query = "SELECT Full_name, orp_one, orp_two FROM Teaching_staff";
                using (OleDbDataAdapter adapter1 = new OleDbDataAdapter(query, conn))
                {
                    adapter1.Fill(teachingStaffTable1);
                }

                int sum2 = 0;
                int sub2 = 0;
                int sub3 = 0;
                int sum3 = 0;

                foreach (DataRow row in teachingStaffTable1.Rows)
                {
                    string teacherName1 = row["Full_name"].ToString();
                    string pro_orp = row["orp_one"].ToString();
                    string pro_orps = row["orp_two"].ToString();
                    int nirm1Value, nirm2Value;
                    

                    if (int.TryParse(pro_orp, out nirm1Value) && int.TryParse(pro_orps, out nirm2Value))
                    {

                        sum2 += nirm1Value;
                        sum3 += nirm2Value;
                        sub2 = sum2 - 2;
                        sub3 = sum3 - 8;
                    }
                    else
                    {
                        // Handle the case where the conversion fails
                        sum2 = 0; // or any other default value you want
                    }

                    DataGridViewRow dataGridViewRow = new DataGridViewRow();
                    dataGridViewRow.CreateCells(dataGridView13, teacherName1, pro_orp,pro_orps);
                    dataGridView13.Rows.Add(dataGridViewRow);
                    dataGridView13.ReadOnly = true;
                }

                dataGridView14.Rows.Add("Всего", sum2.ToString(), sum3.ToString());
                dataGridView14.Rows.Add("Вакансий", sub2.ToString(), sub3.ToString());
                conn.Close();
            }
        }

        private void radioButton24_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void radioButton25_CheckedChanged(object sender, EventArgs e)
        {

        }
















        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView13_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView14_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void listBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView15_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView14_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView16_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView16_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView17_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }

        private void dataGridView18_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // Change the background color of the cell
            e.CellStyle.BackColor = Color.AntiqueWhite;
            // Change the background color of the header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.AntiqueWhite;
        }
        private void dataGridView15_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void dataGridView15_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            // Set the DataGridView's ReadOnly property to false
            dataGridView13.ReadOnly = false;

            // Enable the AllowUserToAddRows and AllowUserToDeleteRows properties if necessary
            dataGridView13.AllowUserToAddRows = true;
            dataGridView13.AllowUserToDeleteRows = true;

            // Enable editing mode for the DataGridView
            dataGridView13.BeginEdit(true);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView13.ReadOnly = true;
            dataGridView13.AllowUserToAddRows = false;
            dataGridView13.AllowUserToDeleteRows = false;
            dataGridView13.BeginEdit(false);
        }
    }





}
