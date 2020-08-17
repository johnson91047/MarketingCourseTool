using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Input;
using MarketingCourseTool.Model;
using MarketingCourseTool.Utility;
using Microsoft.Office.Interop.Excel;
using Prism.Commands;
using Prism.Mvvm;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;
using Window = System.Windows.Window;

namespace MarketingCourseTool.ViewModel
{
    public class MainViewModel : BindableBase
    {
        private const string ExcelTemplatePath = @"./template.xlsx";
        private string _docPath;
        private string _folderPath;
        private string _excelFolderPath;
        private string _excelPath;
        private string _templateMessage;
        private bool _noSameIndex;
        private bool _noSameGroup;
        private bool _noInterchange;
        private bool _twoVersion;
        private bool _canFinishExcel;
        private List<StudentData> _studentData;
        private Dictionary<(int, int), bool> _counter;
        private Window _parent;

        public MainViewModel(Window parent)
        {
            _parent = parent;
            _noSameIndex = true;
            _noSameGroup = true;
            _noInterchange = true;
            _twoVersion = true;
            _canFinishExcel = false;
            _counter = new Dictionary<(int, int), bool>();
            _studentData = new List<StudentData>();
            OpenDocFolderCommand = new DelegateCommand(OpenDocFolder);
            OpenFolderFolderCommand = new DelegateCommand(OpenFolderFolder);
            OpenExcelFolderCommand = new DelegateCommand(OpenExcelFolder);
            CloseAppCommand = new DelegateCommand(System.Windows.Application.Current.Shutdown);
            ClearCommand = new DelegateCommand(ClearAll);
            FinishExcelCommand = new DelegateCommand(FinishExcel, () => _canFinishExcel);
            GenerateCommand = new DelegateCommand(Generate);
        }

        public ICommand CloseAppCommand { get; }

        public ICommand GenerateCommand { get; }

        public ICommand FinishExcelCommand { get; }
        public ICommand ClearCommand { get; }
        public ICommand OpenDocFolderCommand { get; }
        public ICommand OpenFolderFolderCommand { get; }
        public ICommand OpenExcelFolderCommand { get; }

        public string DocumentFilePath
        {
            get => _docPath;
            set => SetProperty(ref _docPath, value);
        }

        public string FolderPath
        {
            get => _folderPath;
            set => SetProperty(ref _folderPath, value);
        }

        public string ExcelPath
        {
            get => _excelFolderPath;
            set => SetProperty(ref _excelFolderPath, value);
        }

        public string TemplateMessage
        {
            get => _templateMessage;
            set => SetProperty(ref _templateMessage, value);
        }

        /*public bool NoSameIndexCheckbox
        {
            get => _noSameIndex;
            set
            {
                if (value)
                {
                    _rule |= GenerateRule.NoSameIndex;
                }
                else
                {
                    _rule &= ~GenerateRule.NoSameIndex;
                }

                SetProperty(ref _noSameIndex, value);
            }
        }

        public bool NoSameGroupCheckbox
        {
            get => _noSameGroup;
            set
            {
                if (value)
                {
                    _rule |= GenerateRule.NoSameGroup;
                }
                else
                {
                    _rule &= ~GenerateRule.NoSameGroup;
                }

                SetProperty(ref _noSameGroup, value);
            }


        }

        public bool NoInterchangeCheckbox
        {
            get => _noInterchange;
            set
            {
                if (value)
                {
                    _rule |= GenerateRule.NoInterchange;
                }
                else
                {
                    _rule &= ~GenerateRule.NoInterchange;
                }

                SetProperty(ref _noInterchange, value);
            }
        }

        public bool TwoVersionCheckbox
        {
            get => _twoVersion;
            set
            {
                if (value)
                {
                    _rule |= GenerateRule.TwoVersion;
                }
                else
                {
                    _rule &= ~GenerateRule.TwoVersion;
                }

                SetProperty(ref _twoVersion, value);
            }
        }*/

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

        private void OpenFolderFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    FolderPath = dialog.SelectedPath;
                }
            }
        }

        private void OpenExcelFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelPath = dialog.SelectedPath;
                }
            }
        }

        private void ClearAll()
        {
            DocumentFilePath = string.Empty;
            FolderPath = string.Empty;
            ExcelPath = string.Empty;
            TemplateMessage = string.Empty;
            _noInterchange = true;
            _noSameGroup = true;
            _noSameIndex = true;
            _twoVersion = true;
            _canFinishExcel = false;
            ((DelegateCommand) FinishExcelCommand).RaiseCanExecuteChanged();
        }

        private void CreateFolderAndIndex()
        {
            int indexCounter = 1;
            DirectoryInfo info = new DirectoryInfo(_docPath);
            string nameChangedFolderPath = Path.Combine(_folderPath, @"編號後的檔案");
            DirectoryInfo nameChangedFolder = Directory.CreateDirectory(nameChangedFolderPath);
            foreach (FileInfo docInfo in info.GetFiles())
            {
                string fileName = Path.GetFileNameWithoutExtension(docInfo.Name);
                string folderPath = Path.Combine(_folderPath, fileName);
                string extension = Path.GetExtension(docInfo.Name);
                if (extension != ".docx" && extension != ".pdf" && extension != ".doc") continue;
                DirectoryInfo createdDir = Directory.CreateDirectory(folderPath);

                docInfo.CopyTo(Path.Combine(createdDir.FullName, docInfo.Name), true);
                docInfo.CopyTo(Path.Combine(nameChangedFolder.FullName, $"{indexCounter}_1{extension}"), true);
                docInfo.CopyTo(Path.Combine(nameChangedFolder.FullName, $"{indexCounter}_2{extension}"), true);

                (string studentId, string name) = SeparateFileName(fileName);
                _studentData.Add(new StudentData
                {
                    StudentId = studentId,
                    Name = name,
                    Index = indexCounter,
                });
                indexCounter++;
            }
        }

        private void ExcelOperation(System.Action<Worksheet> action, bool copyFile = false)
        {
            _excelPath = Path.Combine(_excelFolderPath, "Result.xlsx");
            if (copyFile)
            {
                File.Copy(ExcelTemplatePath, _excelPath, true);
            }

            ExcelApplication excel = new ExcelApplication {Visible = false, DisplayAlerts = false};
            Workbook workbook = excel.Workbooks.Open(Path.GetFullPath(_excelPath));
            Worksheet sheet = workbook.Sheets["main"];
            try
            {
                action(sheet);
                workbook.Save();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Marshal.ReleaseComObject(sheet);
                workbook.Close();
                excel.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excel);
            }

        }

        private void FillExcel(Worksheet sheet)
        {
            int row = 2;
            foreach (StudentData studentData in _studentData)
            {
                sheet.Cells[row, "A"] = studentData.Index;
                sheet.Cells[row, "B"] = $"'{studentData.StudentId}";
                sheet.Cells[row, "C"] = studentData.Name;
                sheet.Cells[row, "D"] = studentData.Group;
                sheet.Cells[row, "E"] = studentData.DocUrl1;
                sheet.Cells[row, "F"] = studentData.DocUrl2;
                sheet.Cells[row, "G"] = studentData.Indexes[0];
                sheet.Cells[row, "H"] = studentData.Indexes[1];
                sheet.Cells[row, "I"] = studentData.Message;

                row++;
            }
        }

        private void ReadFromExcel(Worksheet sheet)
        {
            int row = 2;
            foreach (StudentData studentData in _studentData)
            {
                studentData.Index       = Convert.ToInt32(((Range)sheet.Cells[row, "A"]).Value.ToString());
                studentData.StudentId   = ((Range)sheet.Cells[row, "B"]).Value.ToString();
                studentData.Name        = ((Range)sheet.Cells[row, "C"]).Value.ToString();
                studentData.Group       = Convert.ToInt32(((Range)sheet.Cells[row, "D"]).Value.ToString());
                studentData.DocUrl1     = ((Range)sheet.Cells[row, "E"]).Value?.ToString();
                studentData.DocUrl2     = ((Range)sheet.Cells[row, "F"]).Value?.ToString();
                studentData.Indexes[0]  = Convert.ToInt32(((Range)sheet.Cells[row, "G"]).Value.ToString());
                studentData.Indexes[1]  = Convert.ToInt32(((Range)sheet.Cells[row, "H"]).Value.ToString());
                studentData.Message     = ((Range)sheet.Cells[row, "I"]).Value?.ToString();

                row++;
            }
        }

        private (string, string) SeparateFileName(string fileName)
        {
            string[] name = fileName.Split(new[] {'_'}, StringSplitOptions.RemoveEmptyEntries);
            return (name[0], name[1]);
        }

        private void Generate()
        {
            _studentData.Clear();
            if (!CheckPath()) return;
            CreateFolderAndIndex();
            ExcelOperation(FillExcel, true);
            StartExcel(Path.Combine(_excelFolderPath, "Result.xlsx"));
            _canFinishExcel = true;
            ((DelegateCommand)FinishExcelCommand).RaiseCanExecuteChanged();
            MessageBox.Show("請填入組別於\"Group\"欄位及Google Drive網址於\"Url\"欄位後存檔關閉，再點擊第二步驟", "", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private bool CheckPath()
        {
            if (string.IsNullOrEmpty(_docPath) ||
                string.IsNullOrEmpty(_folderPath) ||
                string.IsNullOrEmpty(_excelFolderPath))
            {
                MessageBox.Show("路徑設定不完整", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void StartExcel(string path)
        {
            Process.Start("excel", path);
        }

        private void FinishExcel()
        {
            
            ExcelOperation(ReadFromExcel);
            GenerateTwoIndex();
            GenerateMessage();
            ExcelOperation(FillExcel);
            StartExcel(Path.Combine(_excelFolderPath, "Result.xlsx"));
            MessageBox.Show("已完成填入編號及訊息", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void GenerateMessage()
        {
            if(string.IsNullOrEmpty(TemplateMessage)) return;
            foreach (StudentData currentStudentData in _studentData)
            {
                string message = TemplateMessage;
                message = ReplaceUtility.Replace(message, @"{編號}", currentStudentData.Index.ToString());
                message = ReplaceUtility.Replace(message, @"{組別}", currentStudentData.Group.ToString());
                message = ReplaceUtility.Replace(message, @"{網址1}", currentStudentData.DocUrl1);
                message = ReplaceUtility.Replace(message, @"{網址2}", currentStudentData.DocUrl2);
                message = ReplaceUtility.Replace(message, @"{編號1}", currentStudentData.Indexes[0].ToString());
                message = ReplaceUtility.Replace(message, @"{編號2}", currentStudentData.Indexes[1].ToString());
                message = ReplaceUtility.Replace(message, @"{編號1網址}",
                    _studentData.FirstOrDefault(s => s.Index == currentStudentData.Indexes[0])?.DocUrl1);
                message = ReplaceUtility.Replace(message, @"{編號2網址}",
                    _studentData.FirstOrDefault(s => s.Index == currentStudentData.Indexes[1])?.DocUrl2);
                currentStudentData.Message = message;
            }
        }

        private void GenerateTwoIndex()
        {
            for (int i = 0; i < _studentData.Count; i++)
            {
                _counter.Add((i + 1, 0), false);
            }
            for (int i = 0; i < _studentData.Count; i++)
            {
                _counter.Add((i + 1, 1), false);
            }
            if (!GenerateIndex())
            {
                MessageBox.Show("無法編號", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            _counter.Clear();
        }

        private bool GenerateIndex()
        {
            bool setFirst = false;
            bool setSecond = false;
            if (!FindUnassignedStudent(out int index)) return true;

            List<int> availableList = AvailableIndex(_studentData[index]);
            for (int i = 0; i < availableList.Count; i++)
            {
                if (_studentData[index].Indexes[0] == 0)
                {
                    if (_counter[(availableList[i], 0)]) continue;
                    _studentData[index].Indexes[0] = availableList[i];
                    setFirst = true;
                    _counter[(availableList[i], 0)] = true;

                }
                else if (_studentData[index].Indexes[1] == 0)
                {
                    if (_counter[(availableList[i], 1)]) continue;
                    _studentData[index].Indexes[1] = availableList[i];
                    setSecond = true;
                    _counter[(availableList[i], 1)] = true;

                }

                if (GenerateIndex())
                {
                    return true;
                }

                if (setFirst)
                {
                    _studentData[index].Indexes[0] = 0;
                    _counter[(availableList[i], 0)] = false;
                }

                if (setSecond)
                {
                    _studentData[index].Indexes[1] = 0;
                    _counter[(availableList[i], 1)] = false;
                }
            }

            return false;
        }

        private bool FindUnassignedStudent(out int index)
        {
            for (int i = 0; i < _studentData.Count; i++)
            {
                if (_studentData[i].Indexes.Contains(0))
                {
                    index = i;
                    return true;
                }
            }
            index = 0;
            return false;
        }

        private List<int> AvailableIndex(StudentData student)
        {
            return _studentData.Where(s => !s.Indexes.Contains(student.Index) &&
                                           !student.Indexes.Contains(s.Index) &&
                                           s.Index != student.Index &&
                                           s.Group != student.Group )
                                .Select(s => s.Index).ToList();
        }

    }
}
