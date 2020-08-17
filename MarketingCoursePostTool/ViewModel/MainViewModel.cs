using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Input;
using Microsoft.Office.Interop.Excel;
using Prism.Commands;
using Prism.Mvvm;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace MarketingCoursePostTool.ViewModel
{
    public class Student
    {
        public string Id;
        public string Name;
    }

    public class MainViewModel : BindableBase
    {
        private const string ReturnFeedbackFolderName = @"發還feedback";
        private string _docPath;
        private string _excelPath;
        private Dictionary<int, Student> _studentDictionary;

        public MainViewModel()
        {
            OpenDocFolderCommand = new DelegateCommand(OpenDocFolder);
            CloseAppCommand = new DelegateCommand(System.Windows.Application.Current.Shutdown);
            GenerateCommand = new DelegateCommand(Generate);
            OpenExcelCommand = new DelegateCommand(OpenExcel);
            _studentDictionary = new Dictionary<int, Student>();
        }

        public ICommand CloseAppCommand { get; }

        public ICommand GenerateCommand { get; }

        public ICommand OpenDocFolderCommand { get; }

        public ICommand OpenExcelCommand { get; }

        public string DocumentFilePath
        {
            get => _docPath;
            set => SetProperty(ref _docPath, value);
        }
        public string ExcelPath
        {
            get => _excelPath;
            set => SetProperty(ref _excelPath, value);
        }

        private void OpenDocFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    DocumentFilePath = dialog.SelectedPath;
                }
            }
        }

        private void OpenExcel()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelPath = dialog.FileName;
                }
            }
        }

        private bool ReadExcel()
        {

            Application excel = new Application { Visible = false, DisplayAlerts = false };
            Workbook workbook = excel.Workbooks.Open(Path.GetFullPath(_excelPath));
            Worksheet sheet = workbook.Sheets["main"];
            Console.WriteLine(sheet.UsedRange.Rows.Count);
            try
            {
                for (int i = 2; i <= sheet.UsedRange.Rows.Count; i++)
                {
                    string id = ((Range)sheet.Cells[i, "B"]).Value.ToString();
                    string name = ((Range)sheet.Cells[i, "C"]).Value.ToString();
                    _studentDictionary.Add(Convert.ToInt32(((Range)sheet.Cells[i, "A"]).Value.ToString()), new Student
                    {
                        Id = id,
                        Name = name
                    });
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _studentDictionary.Clear();
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);
            }

            return true;
        }

        private (string otherStudentIndex, string version, string id, string name) DeconstructFileName(string fileName)
        {
            string[] array = fileName.Split('_');
            if(array.Length != 5) throw new ArgumentException($"文件檔名有誤 => {fileName}");
            return (array[0], array[1], array[2], array[3]);
        }

        private string IndexToIdName(int index)
        {
            if (_studentDictionary.TryGetValue(index, out Student student))
            {
                return $"{index}_{student.Id}_{student.Name}";
            }

            throw new DirectoryNotFoundException($"Index錯誤，找不到Index: {index}");
        }

        private string IndexToIdNameWithoutIndex(int index)
        {
            if (_studentDictionary.TryGetValue(index, out Student student))
            {
                return $"{student.Id}_{student.Name}";
            }

            throw new DirectoryNotFoundException($"Index錯誤，找不到Index: {index}");
        }

        private void Generate()
        {
            if (!CheckPath()) return;
            if(!ReadExcel()) return;
            string returnFeedbackFolderPath = Path.Combine(_docPath, ReturnFeedbackFolderName);
            DirectoryInfo info = new DirectoryInfo(_docPath);
            try
            {
                if (!Directory.Exists(returnFeedbackFolderPath))
                {
                    Directory.CreateDirectory(returnFeedbackFolderPath);
                }

                foreach (FileInfo docInfo in info.GetFiles())
                {
                    string fileName = Path.GetFileNameWithoutExtension(docInfo.Name);
                    string extension = Path.GetExtension(docInfo.Name);

                    if(extension != ".docx" && extension != ".pdf" && extension != ".doc" ) continue;

                    (string otherStudentIndex, string version, string id, string name) = DeconstructFileName(fileName);

                    int currentIndex = _studentDictionary.FirstOrDefault(p => p.Value.Id == id).Key;

                    // to feedback giver folder
                    string folderPath = Path.Combine(_docPath, IndexToIdName(currentIndex));
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }


                    string feedbackFolderPath = Path.Combine(folderPath, $"{id}_{name}_feedback");
                    if (!Directory.Exists(feedbackFolderPath))
                    {
                        Directory.CreateDirectory(feedbackFolderPath);
                    }

                    docInfo.CopyTo(Path.Combine(feedbackFolderPath, docInfo.Name), true);


                    // to feedback receiver folder
                    string receiverFolderName = IndexToIdName(Convert.ToInt32(otherStudentIndex));

                    folderPath = Path.Combine(_docPath, receiverFolderName);
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    string feedbackReceiverPath = Path.Combine(folderPath, $"feedback_from_{id}_{name}");
                    if (!Directory.Exists(feedbackReceiverPath))
                    {
                        Directory.CreateDirectory(feedbackReceiverPath);
                    }

                    docInfo.CopyTo(Path.Combine(feedbackReceiverPath, docInfo.Name), true);


                    // to return feedback folder
                    folderPath = Path.Combine(returnFeedbackFolderPath, $"feedback_{IndexToIdNameWithoutIndex(Convert.ToInt32(otherStudentIndex))}");
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    docInfo.CopyTo(Path.Combine(folderPath, $"feedback{version}_{IndexToIdNameWithoutIndex(Convert.ToInt32(otherStudentIndex))}{extension}"), true);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Process.Start(_docPath);
            MessageBox.Show("產生完成", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private bool CheckPath()
        {
            if (string.IsNullOrEmpty(_docPath) ||
                string.IsNullOrEmpty(_excelPath))
            {
                MessageBox.Show("路徑設定不完整", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
    }
}
