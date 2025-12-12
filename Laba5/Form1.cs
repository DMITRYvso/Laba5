using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.Validation;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Laba5
{
    public partial class Form1 : Form
    {
        Laba5Entities1 conn = new Laba5Entities1();

        private int maxId = 0;
        private List<int> existingIds = new List<int>();

        public Form1()
        {
            InitializeComponent();
            GetIdsFromDatabase();
            SetupDataGridView();
            LoadData();
            AddEmptyRow();
        }

        private void GetIdsFromDatabase()
        {
            try
            {
                if (conn.Rabota4.Any())
                {
                    maxId = conn.Rabota4.Max(r => r.ID);
                    existingIds = conn.Rabota4.Select(r => r.ID).ToList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении данных из БД: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SetupDataGridView()
        {
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = true;
            dataGridView1.AutoGenerateColumns = false;

            if (dataGridView1.Columns.Count == 0)
            {
                dataGridView1.Columns.Add("ID", "ID");
                dataGridView1.Columns.Add("Name", "ФИО");
                dataGridView1.Columns.Add("Age", "Возраст");
                dataGridView1.Columns.Add("Department", "Отдел");
                dataGridView1.Columns.Add("Salary", "Зарплата");
                dataGridView1.Columns.Add("Date_of_admission", "Дата приема");
                dataGridView1.Columns["Date_of_admission"].ReadOnly = true;
            }
        }

        private void LoadData()
        {
            try
            {
                var data = conn.Rabota4.ToList();
                dataGridView1.Rows.Clear();
                foreach (var item in data)
                {
                    dataGridView1.Rows.Add(
                        item.ID,
                        item.Name,
                        item.Age,
                        item.Department,
                        item.Salary,
                        item.Date_of_admission
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddEmptyRow()
        {
            int rowIndex = dataGridView1.Rows.Add();
            int newId = maxId + 1;
            dataGridView1.Rows[rowIndex].Cells["ID"].Value = newId;
            maxId = newId;

            dataGridView1.Rows[rowIndex].Cells["Date_of_admission"].Value = DateTime.Today;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                List<Rabota4> recordsToAdd = new List<Rabota4>();
                List<int> processedIds = new List<int>();
                bool hasErrors = false;
                string errorMessage = "";

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (IsRowEmpty(row))
                        continue;

                    int id;
                    if (row.Cells["ID"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["ID"].Value.ToString()))
                    {
                        if (!int.TryParse(row.Cells["ID"].Value.ToString(), out id))
                        {
                            hasErrors = true;
                            errorMessage = $"Неверный формат ID (строка {row.Index + 1})";
                            break;
                        }
                    }
                    else
                    {
                        hasErrors = true;
                        errorMessage = $"ID не может быть пустым (строка {row.Index + 1})";
                        break;
                    }

                    if (existingIds.Contains(id))
                    {
                        continue;
                    }

                    Rabota4 newRecord = new Rabota4();

                    if (id <= 0)
                    {
                        hasErrors = true;
                        errorMessage = $"ID должен быть больше 0 (строка {row.Index + 1})";
                        break;
                    }

                    if (processedIds.Contains(id))
                    {
                        hasErrors = true;
                        errorMessage = $"ID {id} повторяется в пределах добавляемых записей (строка {row.Index + 1})";
                        break;
                    }

                    if (existingIds.Contains(id))
                    {
                        hasErrors = true;
                        errorMessage = $"ID {id} уже существует в базе данных (строка {row.Index + 1})";
                        break;
                    }

                    newRecord.ID = id;
                    processedIds.Add(id);

                    if (id > maxId) maxId = id;

                    if (row.Cells["Name"].Value != null)
                    {
                        string name = row.Cells["Name"].Value.ToString();
                        if (string.IsNullOrWhiteSpace(name))
                        {
                            hasErrors = true;
                            errorMessage = $"ФИО не может быть пустым (строка {row.Index + 1})";
                            break;
                        }

                        if (name.Length > 100)
                        {
                            hasErrors = true;
                            errorMessage = $"ФИО не может превышать 100 символов (строка {row.Index + 1})";
                            break;
                        }
                        newRecord.Name = name;
                    }
                    else
                    {
                        hasErrors = true;
                        errorMessage = $"ФИО не может быть пустым (строка {row.Index + 1})";
                        break;
                    }

                    if (row.Cells["Age"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["Age"].Value.ToString()))
                    {
                        int age;
                        if (int.TryParse(row.Cells["Age"].Value.ToString(), out age))
                        {
                            if (age < 18 || age > 65)
                            {
                                hasErrors = true;
                                errorMessage = $"Возраст должен быть от 18 до 65 лет (строка {row.Index + 1})";
                                break;
                            }
                            newRecord.Age = age;
                        }
                        else
                        {
                            hasErrors = true;
                            errorMessage = $"Возраст должен содержать только цифры (строка {row.Index + 1})";
                            break;
                        }
                    }
                    else
                    {
                        hasErrors = true;
                        errorMessage = $"Возраст не может быть пустым (строка {row.Index + 1})";
                        break;
                    }

                    if (row.Cells["Department"].Value != null)
                    {
                        string department = row.Cells["Department"].Value.ToString();
                        if (string.IsNullOrWhiteSpace(department))
                        {
                            hasErrors = true;
                            errorMessage = $"Отдел не может быть пустым (строка {row.Index + 1})";
                            break;
                        }

                        if (department.Length > 100)
                        {
                            hasErrors = true;
                            errorMessage = $"Название отдела не может превышать 100 символов (строка {row.Index + 1})";
                            break;
                        }
                        newRecord.Department = department;
                    }
                    else
                    {
                        hasErrors = true;
                        errorMessage = $"Отдел не может быть пустым (строка {row.Index + 1})";
                        break;
                    }

                    if (row.Cells["Salary"].Value != null && !string.IsNullOrWhiteSpace(row.Cells["Salary"].Value.ToString()))
                    {
                        decimal salary;
                        if (decimal.TryParse(row.Cells["Salary"].Value.ToString(), out salary))
                        {
                            if (salary < 10000 || salary > 1000000)
                            {
                                hasErrors = true;
                                errorMessage = $"Зарплата должна быть от 10000 до 1000000 (строка {row.Index + 1})";
                                break;
                            }
                            newRecord.Salary = salary;
                        }
                        else
                        {
                            hasErrors = true;
                            errorMessage = $"Неверный формат зарплаты (строка {row.Index + 1})";
                            break;
                        }
                    }
                    else
                    {
                        hasErrors = true;
                        errorMessage = $"Зарплата не может быть пустой (строка {row.Index + 1})";
                        break;
                    }

                    if (row.Cells["Date_of_admission"].Value != null)
                    {
                        newRecord.Date_of_admission = (DateTime)row.Cells["Date_of_admission"].Value;
                    }
                    else
                    {
                        newRecord.Date_of_admission = DateTime.Today;
                    }

                    recordsToAdd.Add(newRecord);
                }

                if (hasErrors)
                {
                    MessageBox.Show(errorMessage, "Ошибка валидации", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (recordsToAdd.Count == 0)
                {
                    MessageBox.Show("Нет новых данных для сохранения. Добавьте новые записи в пустую строку.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                foreach (var record in recordsToAdd)
                {
                    conn.Rabota4.Add(record);
                    existingIds.Add(record.ID);
                }

                conn.SaveChanges();

                MessageBox.Show($"Данные успешно сохранены! Добавлено {recordsToAdd.Count} записей.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadData();
                AddEmptyRow();
            }
            catch (DbEntityValidationException ex)
            {
                foreach (var validationErrors in ex.EntityValidationErrors)
                {
                    foreach (var validationError in validationErrors.ValidationErrors)
                    {
                        MessageBox.Show($"Свойство: {validationError.PropertyName} Ошибка: {validationError.ErrorMessage}", "Ошибка валидации БД", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool IsRowEmpty(DataGridViewRow row)
        {
            bool nameIsEmpty = row.Cells["Name"].Value == null || string.IsNullOrWhiteSpace(row.Cells["Name"].Value.ToString());
            bool ageIsEmpty = row.Cells["Age"].Value == null || string.IsNullOrWhiteSpace(row.Cells["Age"].Value.ToString());
            bool deptIsEmpty = row.Cells["Department"].Value == null || string.IsNullOrWhiteSpace(row.Cells["Department"].Value.ToString());
            bool salaryIsEmpty = row.Cells["Salary"].Value == null || string.IsNullOrWhiteSpace(row.Cells["Salary"].Value.ToString());
            return nameIsEmpty && ageIsEmpty && deptIsEmpty && salaryIsEmpty;
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "Age")
            {
                if (!string.IsNullOrEmpty(e.FormattedValue.ToString()))
                {
                    if (!Regex.IsMatch(e.FormattedValue.ToString(), @"^\d+$"))
                    {
                        dataGridView1.Rows[e.RowIndex].ErrorText = "Возраст должен содержать только цифры";
                        e.Cancel = true;
                    }
                    else
                    {
                        dataGridView1.Rows[e.RowIndex].ErrorText = "";
                    }
                }
            }

            if (dataGridView1.Columns[e.ColumnIndex].Name == "Salary")
            {
                if (!string.IsNullOrEmpty(e.FormattedValue.ToString()))
                {
                    if (!Regex.IsMatch(e.FormattedValue.ToString(), @"^\d+([,.]\d+)?$"))
                    {
                        dataGridView1.Rows[e.RowIndex].ErrorText = "Зарплата должна быть числом";
                        e.Cancel = true;
                    }
                    else
                    {
                        dataGridView1.Rows[e.RowIndex].ErrorText = "";
                    }
                }
            }

            if (dataGridView1.Columns[e.ColumnIndex].Name == "ID")
            {
                if (!string.IsNullOrEmpty(e.FormattedValue.ToString()))
                {
                    if (!Regex.IsMatch(e.FormattedValue.ToString(), @"^\d+$"))
                    {
                        dataGridView1.Rows[e.RowIndex].ErrorText = "ID должен содержать только цифры";
                        e.Cancel = true;
                    }
                    else
                    {
                        dataGridView1.Rows[e.RowIndex].ErrorText = "";
                    }
                }
            }
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count)
            {
                var row = dataGridView1.Rows[e.RowIndex];
                if (row.Cells["Date_of_admission"].Value == null)
                {
                    row.Cells["Date_of_admission"].Value = DateTime.Today;
                }
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "ID")
            {
                var cell = dataGridView1.Rows[e.RowIndex].Cells["ID"];
                if (cell.Value != null && !string.IsNullOrWhiteSpace(cell.Value.ToString()))
                {
                    int id;
                    if (int.TryParse(cell.Value.ToString(), out id))
                    {
                        if (existingIds.Contains(id))
                        {
                            MessageBox.Show($"ID {id} уже существует в базе данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            int newId = maxId + 1;
                            cell.Value = newId;
                            maxId = newId;
                        }
                        else
                        {
                            bool isDuplicate = false;
                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                if (i != e.RowIndex && !dataGridView1.Rows[i].IsNewRow)
                                {
                                    var otherCell = dataGridView1.Rows[i].Cells["ID"];
                                    if (otherCell.Value != null && otherCell.Value.ToString() == cell.Value.ToString())
                                    {
                                        isDuplicate = true;
                                        break;
                                    }
                                }
                            }

                            if (isDuplicate)
                            {
                                MessageBox.Show($"ID {id} уже используется в другой строке", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                int newId = maxId + 1;
                                cell.Value = newId;
                                maxId = newId;
                            }
                            else if (id > maxId)
                            {
                                maxId = id;
                            }
                        }
                    }
                }
            }
        }

        private void btnAddRow_Click(object sender, EventArgs e)
        {
            AddEmptyRow();
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows.Count - 1;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells["Name"];
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (dataGridView1.CurrentCell != null)
                {
                    int currentRow = dataGridView1.CurrentCell.RowIndex;
                    int currentCol = dataGridView1.CurrentCell.ColumnIndex;

                    if (currentCol == dataGridView1.Columns.Count - 1)
                    {
                        if (currentRow == dataGridView1.Rows.Count - 1)
                        {
                            AddEmptyRow();
                            dataGridView1.CurrentCell = dataGridView1.Rows[currentRow + 1].Cells[0];
                            e.Handled = true;
                        }
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (!conn.Rabota4.Any())
                {
                    MessageBox.Show("В базе данных нет записей.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                decimal maxSalary = (decimal)conn.Rabota4.Max(r => r.Salary);
                var topEmployees = conn.Rabota4
                    .Where(r => r.Salary == maxSalary)
                    .ToList();

                if (topEmployees.Any())
                {
                    string message = "Сотрудник(и) с самой высокой зарплатой:\n\n";

                    foreach (var employee in topEmployees)
                    {
                        message += $"ФИО: {employee.Name}\n" +
                                  $"Зарплата: {employee.Salary:N2} руб.\n" +
                                  $"Отдел: {employee.Department}\n" +
                                  $"Возраст: {employee.Age} лет\n" +
                                  $"----------------------------------\n";
                    }

                    MessageBox.Show(message, "Самые высокооплачиваемые сотрудники",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Не удалось найти сотрудников.",
                        "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении данных: {ex.Message}",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0 && dataGridView1.SelectedCells.Count == 0)
                {
                    MessageBox.Show("Пожалуйста, выберите одну или несколько строк для удаления.",
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                List<DataGridViewRow> rowsToDelete = new List<DataGridViewRow>();

                if (dataGridView1.SelectedRows.Count > 0)
                {
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        if (!row.IsNewRow)
                        {
                            rowsToDelete.Add(row);
                        }
                    }
                }
                else if (dataGridView1.SelectedCells.Count > 0)
                {
                    HashSet<int> rowIndices = new HashSet<int>();

                    foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                    {
                        if (cell.RowIndex >= 0 && !dataGridView1.Rows[cell.RowIndex].IsNewRow)
                        {
                            rowIndices.Add(cell.RowIndex);
                        }
                    }

                    foreach (int rowIndex in rowIndices)
                    {
                        rowsToDelete.Add(dataGridView1.Rows[rowIndex]);
                    }
                }

                if (rowsToDelete.Count == 0)
                {
                    MessageBox.Show("Нет выбранных строк для удаления. Убедитесь, что вы не выбрали пустую строку для ввода.",
                        "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                List<int> idsToDeleteFromDb = new List<int>();
                List<string> employeeNames = new List<string>();

                foreach (DataGridViewRow row in rowsToDelete)
                {
                    if (row.Cells["ID"].Value != null &&
                        int.TryParse(row.Cells["ID"].Value.ToString(), out int id))
                    {
                        idsToDeleteFromDb.Add(id);
                    }

                    string name = row.Cells["Name"].Value?.ToString() ?? "Неизвестный";
                    employeeNames.Add(name);
                }

                string confirmationMessage;

                if (rowsToDelete.Count == 1)
                {
                    confirmationMessage = $"Вы уверены, что хотите удалить сотрудника:\n\n" +
                                        $"ФИО: {employeeNames[0]}\n" +
                                        $"ID: {(idsToDeleteFromDb.Count > 0 ? idsToDeleteFromDb[0].ToString() : "Нет ID")}\n\n" +
                                        $"Это действие нельзя отменить!";
                }
                else
                {
                    confirmationMessage = $"Вы уверены, что хотите удалить {rowsToDelete.Count} записей?\n\n";

                    int namesToShow = Math.Min(3, employeeNames.Count);
                    for (int i = 0; i < namesToShow; i++)
                    {
                        confirmationMessage += $"• {employeeNames[i]}\n";
                    }

                    if (employeeNames.Count > 3)
                    {
                        confirmationMessage += $"... и еще {employeeNames.Count - 3} записей\n";
                    }

                    confirmationMessage += "\nЭто действие нельзя отменить!";
                }

                DialogResult result = MessageBox.Show(
                    confirmationMessage,
                    "Подтверждение удаления",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2);

                if (result == DialogResult.Yes)
                {
                    int deletedFromDbCount = 0;
                    if (idsToDeleteFromDb.Count > 0)
                    {
                        foreach (int id in idsToDeleteFromDb)
                        {
                            var recordToDelete = conn.Rabota4.FirstOrDefault(r => r.ID == id);
                            if (recordToDelete != null)
                            {
                                conn.Rabota4.Remove(recordToDelete);
                                deletedFromDbCount++;
                                existingIds.Remove(id);
                            }
                        }

                        if (deletedFromDbCount > 0)
                        {
                            conn.SaveChanges();
                        }
                    }

                    rowsToDelete = rowsToDelete.OrderByDescending(r => r.Index).ToList();

                    foreach (DataGridViewRow row in rowsToDelete)
                    {
                        dataGridView1.Rows.Remove(row);
                    }

                    bool hasDataRows = false;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (!row.IsNewRow && !IsRowEmpty(row))
                        {
                            hasDataRows = true;
                            break;
                        }
                    }

                    if (!hasDataRows || dataGridView1.Rows.Count == 0)
                    {
                        dataGridView1.Rows.Clear();
                        AddEmptyRow();
                    }
                    else
                    {
                        bool hasEmptyRow = false;
                        if (dataGridView1.Rows.Count > 0)
                        {
                            var lastRow = dataGridView1.Rows[dataGridView1.Rows.Count - 1];
                            if (lastRow.IsNewRow || IsRowEmpty(lastRow))
                            {
                                hasEmptyRow = true;
                            }
                        }

                        if (!hasEmptyRow)
                        {
                            AddEmptyRow();
                        }
                    }

                    string successMessage;
                    if (deletedFromDbCount > 0)
                    {
                        successMessage = $"Успешно удалено {rowsToDelete.Count} записей " +
                                       $"({deletedFromDbCount} из базы данных).";
                    }
                    else
                    {
                        successMessage = $"Успешно удалено {rowsToDelete.Count} записей из таблицы.";
                    }

                    MessageBox.Show(successMessage, "Успех",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string searchText = textBox1.Text?.Trim();

                if (string.IsNullOrWhiteSpace(searchText))
                {
                    MessageBox.Show("Введите ФИО для поиска.", "Информация",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                dataGridView1.ClearSelection();

                int foundCount = 0;
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    var row = dataGridView1.Rows[i];

                    if (row.IsNewRow) continue;

                    if (row.Cells["Name"].Value != null)
                    {
                        string name = row.Cells["Name"].Value.ToString();

                        if (name.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            row.Selected = true;
                            foundCount++;
                        }
                    }
                }

                if (foundCount > 0)
                {
                    var firstSelected = dataGridView1.SelectedRows[0];
                    dataGridView1.FirstDisplayedScrollingRowIndex = firstSelected.Index;

                    MessageBox.Show($"Найдено {foundCount} записей. Строки выделены в таблице.",
                        "Результаты поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Сотрудники по запросу '{searchText}' не найдены.",
                        "Результаты поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows.Count == 0 ||
                    (dataGridView1.Rows.Count == 1 && dataGridView1.Rows[0].IsNewRow))
                {
                    MessageBox.Show("Нет данных для сохранения.", "Информация",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Текстовый файл (*.txt)|*.txt";
                saveDialog.Title = "Сохранить данные сотрудников";
                saveDialog.FileName = "сотрудники.txt";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    using (StreamWriter sw = new StreamWriter(saveDialog.FileName, false, Encoding.UTF8))
                    {
                        string header = string.Format("{0,-5} {1,-25} {2,-8} {3,-15} {4,-12} {5,-12}",
                            "ID", "ФИО", "Возраст", "Отдел", "Зарплата", "Дата приема");

                        sw.WriteLine("СПИСОК СОТРУДНИКОВ");
                        sw.WriteLine(new string('=', header.Length));
                        sw.WriteLine(header);
                        sw.WriteLine(new string('-', header.Length));

                        int count = 0;
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.IsNewRow || IsRowEmpty(row))
                                continue;

                            string id = row.Cells["ID"].Value?.ToString() ?? "";
                            string name = row.Cells["Name"].Value?.ToString() ?? "";
                            string age = row.Cells["Age"].Value?.ToString() ?? "";
                            string dept = row.Cells["Department"].Value?.ToString() ?? "";
                            string salary = row.Cells["Salary"].Value?.ToString() ?? "";
                            string date = row.Cells["Date_of_admission"].Value?.ToString() ?? "";

                            if (decimal.TryParse(salary, out decimal sal))
                                salary = sal.ToString("N2") + " руб.";

                            if (DateTime.TryParse(date, out DateTime dt))
                                date = dt.ToString("dd.MM.yyyy");

                            string line = string.Format("{0,-5} {1,-25} {2,-8} {3,-15} {4,-12} {5,-12}",
                                id, name, age, dept, salary, date);

                            sw.WriteLine(line);
                            count++;
                        }

                        sw.WriteLine(new string('-', header.Length));
                        sw.WriteLine($"Всего записей: {count}");
                        sw.WriteLine($"Дата экспорта: {DateTime.Now:dd.MM.yyyy HH:mm}");
                    }

                    if (MessageBox.Show("Файл успешно сохранен. Открыть его?", "Сохранение завершено",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start("notepad.exe", saveDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt|CSV файлы (*.csv)|*.csv|Все файлы (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Title = "Выберите файл для импорта данных";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;

                    if (!File.Exists(filePath))
                    {
                        MessageBox.Show("Выбранный файл не существует.", "Ошибка",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    ImportDataFromFileSimple(filePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте данных: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportDataFromFileSimple(string filePath)
        {
            try
            {
                List<Rabota4> validRecords = new List<Rabota4>();
                List<string> errorMessages = new List<string>();
                int lineNumber = 0;

                string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);

                foreach (string line in lines)
                {
                    lineNumber++;

                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    if (lineNumber == 1 && (line.Contains("ID") || line.Contains("ФИО")))
                        continue;

                    Rabota4 record = ParseSimpleLine(line, lineNumber);

                    if (record != null)
                    {
                        if (IsRecordValid(record, lineNumber))
                        {
                            validRecords.Add(record);
                        }
                    }
                }

                if (validRecords.Count == 0)
                {
                    MessageBox.Show("В файле не найдено корректных записей для импорта.", "Информация",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DialogResult result = MessageBox.Show(
                    $"Найдено {validRecords.Count} корректных записей для импорта.\n\nПродолжить?",
                    "Подтверждение импорта",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    int importedCount = 0;

                    foreach (var record in validRecords)
                    {
                        try
                        {
                            if (!conn.Rabota4.Any(r => r.ID == record.ID))
                            {
                                conn.Rabota4.Add(record);
                                conn.SaveChanges();
                                importedCount++;

                                existingIds.Add(record.ID);
                                if (record.ID > maxId)
                                    maxId = record.ID;
                            }
                            else
                            {
                                errorMessages.Add($"ID {record.ID} уже существует в БД");
                            }
                        }
                        catch (Exception ex)
                        {
                            errorMessages.Add($"Ошибка при импорте ID {record.ID}: {ex.Message}");
                        }
                    }

                    LoadData();
                    AddEmptyRow();

                    string message = $"Импортировано {importedCount} записей.";
                    if (errorMessages.Count > 0)
                    {
                        message += $"\n\nОшибки:\n{string.Join("\n", errorMessages.Take(5))}";
                        if (errorMessages.Count > 5)
                            message += $"\n... и еще {errorMessages.Count - 5} ошибок";
                    }

                    MessageBox.Show(message, "Результат импорта",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте файла: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Rabota4 ParseSimpleLine(string line, int lineNumber)
        {
            try
            {
                char[] delimiters = { ';', '\t', ',' };
                string[] parts = line.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length < 4)
                {
                    return null;
                }

                Rabota4 record = new Rabota4();

                if (int.TryParse(parts[0].Trim(), out int id))
                    record.ID = id;
                else
                    return null;

                record.Name = parts[1].Trim();

                if (parts.Length > 2 && int.TryParse(parts[2].Trim(), out int age))
                    record.Age = age;
                else
                    record.Age = 25;

                record.Department = parts[3].Trim();

                if (parts.Length > 4 && decimal.TryParse(parts[4].Trim(), out decimal salary))
                    record.Salary = salary;
                else
                    record.Salary = 30000;

                if (parts.Length > 5 && DateTime.TryParse(parts[5].Trim(), out DateTime date))
                    record.Date_of_admission = date;
                else
                    record.Date_of_admission = DateTime.Today;

                return record;
            }
            catch
            {
                return null;
            }
        }

        private bool IsRecordValid(Rabota4 record, int lineNumber)
        {
            if (record.ID <= 0)
                return false;

            if (string.IsNullOrWhiteSpace(record.Name) || record.Name.Length > 100)
                return false;

            if (record.Age < 18 || record.Age > 65)
                return false;

            if (string.IsNullOrWhiteSpace(record.Department) || record.Department.Length > 100)
                return false;

            if (record.Salary < 10000 || record.Salary > 1000000)
                return false;

            return true;
        }

        private void просмотрToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1.Text = "Сохранение БД отключено";
            button2.Text = "Удаление отключено";
            button4.Text = "Эспорт отключено";
            button5.Text = "Сохраниние txt отключено";
            button1.Enabled = false;
            button2.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button3.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;

        }

        private void полныйДоступToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1.Text = "Сохранить БД";
            button2.Text = "Удалить";
            button4.Text = "Эспорт txt";
            button5.Text = "Сохранить txt";
            button1.Enabled = true;
            button2.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button3.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
        }
    }
}