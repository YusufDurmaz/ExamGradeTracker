using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MahApps.Metro.Controls;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace ExamGradeTracker
{
    public partial class MainWindow : MetroWindow
    {
        public ObservableCollection<Student> Students { get; set; }

        private readonly string _filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExamGradeTracker",
            "students.json"
        );

        public MainWindow()
        {
            InitializeComponent();
            Students = new ObservableCollection<Student>();
            LoadStudents();
            StudentsGrid.ItemsSource = Students;
        }

        private void AddStudent_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(txtStudentName.Text))
            {
                Students.Add(new Student(txtStudentName.Text, Array.Empty<int>()));
                txtStudentName.Clear();
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir öğrenci adı girin.", "Hata", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void StudentsGrid_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (StudentsGrid.SelectedItem is Student selectedStudent)
            {
                ContextMenu contextMenu = new ContextMenu();
                MenuItem deleteMenuItem = new() { Header = $"Delete {selectedStudent.StudentName}" };
                deleteMenuItem.Click += (s, args) => Students.Remove(selectedStudent);
                contextMenu.Items.Add(deleteMenuItem);
                StudentsGrid.ContextMenu = contextMenu;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void SaveStudents()
        {
            var directory = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var json = JsonConvert.SerializeObject(Students);
            File.WriteAllText(_filePath, json);
        }

        private void LoadStudents()
        {
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath);
                var students = JsonConvert.DeserializeObject<ObservableCollection<Student>>(json);
                Students = students ?? new ObservableCollection<Student>();
            }
            else
            {
                Students = new ObservableCollection<Student>();
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            SaveStudents();
            base.OnClosed(e);
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Save an Excel File"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Students");

                    worksheet.Cells[1, 1].Value = "Student Name";
                    worksheet.Cells[1, 2].Value = "Average";
                    worksheet.Cells[1, 3].Value = "Grades";

                    for (int i = 0; i < Students.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = Students[i].StudentName;
                        worksheet.Cells[i + 2, 2].Value = Students[i].Avarage;
                        worksheet.Cells[i + 2, 3].Value = Students[i].GradesString;
                    }

                    var file = new FileInfo(saveFileDialog.FileName);
                    package.SaveAs(file);
                }

                MessageBox.Show("Students exported to Excel successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        private void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Open an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var file = new FileInfo(openFileDialog.FileName);
                using (var package = new ExcelPackage(file))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var student = new Student(
                            worksheet.Cells[row, 1].Text,
                            worksheet.Cells[row, 3].Text.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                                        .Select(int.Parse)
                                                        .ToArray()
                        );
                        Students.Add(student);
                    }
                }

                MessageBox.Show("Students imported from Excel successfully.", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }


    public class Student : INotifyPropertyChanged
    {
        private string studentName;
        private double avarage;
        private int[] grades;
        private string gradesString;

        public string StudentName
        {
            get => studentName;
            set
            {
                studentName = value;
                OnPropertyChanged(nameof(StudentName));
            }
        }

        public double Avarage
        {
            get => avarage;
            set
            {
                avarage = value;
                OnPropertyChanged(nameof(Avarage));
            }
        }

        public int[] Grades
        {
            get => grades;
            set
            {
                grades = value;
                Avarage = GetAverageGrade(grades);
                OnPropertyChanged(nameof(Grades));
                OnPropertyChanged(nameof(GradesString));
            }
        }

        public string GradesString
        {
            get => gradesString;
            set
            {
                gradesString = value;
                if (!string.IsNullOrEmpty(value) && !value.EndsWith(","))
                {
                    try
                    {
                        Grades = value.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(int.Parse)
                                      .ToArray();
                        Avarage = GetAverageGrade(Grades);
                        OnPropertyChanged(nameof(GradesString));
                    }
                    catch
                    {
                        // Handle parsing errors if necessary
                    }
                }
                OnPropertyChanged(nameof(GradesString));
            }
        }

        public Student(string studentName, int[] grades)
        {
            StudentName = studentName;
            Grades = grades;
            Avarage = GetAverageGrade(grades);
            GradesString = string.Join(", ", grades);
        }

        public double GetAverageGrade(int[] grades)
        {
            if (grades.Length == 0)
                return 0;
            return grades.Average();
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
