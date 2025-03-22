using System;
using System.Data;
using System.Reflection;
using Microsoft.WindowsAPICodePack.Dialogs;
using ExcelDataReader;
using System.Diagnostics;
using System.Text;
using Newtonsoft.Json.Linq;

namespace ExcelParser
{
    public partial class Form_ExcelParser : Form
    {
        private const string PATH_SAVED = @".\Path.txt";
        private const string PATH_CLASS_TEMPLATE = @".\ClassTemplate.txt";

        private Dictionary<string, FieldInfo> fieldByName = new Dictionary<string, FieldInfo>();
        private List<DataTableCollection> excelList = new List<DataTableCollection>();

        private string classTemplate;

        private object lockObj = new object();

        #region Init
        public Form_ExcelParser()
        {
            InitializeComponent();
            RegisterEncoding();
            CacheFields();

            LoadSavedPaths();
            LoadClassTemplate();
        }

        private void RegisterEncoding()
        {
            // .Net 5.0 �̻���� ����� ExcelDataReader�� ���ڵ� ���� ������
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void CacheFields()
        {
            var fields = typeof(Form_ExcelParser).GetFields(BindingFlags.Instance | BindingFlags.NonPublic);
            foreach (var f in fields) fieldByName.Add(f.Name, f);
        }

        private void LoadSavedPaths()
        {
            // ������ ���� ���θ� üũ�Ѵ�. ���ٸ� ������ �����Ѵ�.
            if (!File.Exists(PATH_SAVED))
            {
                File.Create(PATH_SAVED).Dispose();
                return;
            }

            // ������ �����Ѵٸ� �о�� dict�� �����Ѵ�.
            var sr = new StreamReader(PATH_SAVED);
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (line == null) continue;
                var words = line.Split('=');

                var tbName = words[0];
                var path = words[1];
                if (fieldByName.ContainsKey(tbName))
                {
                    var textBox = fieldByName[tbName].GetValue(this) as TextBox;
                    if (textBox == null) continue;
                    textBox.Text = path;
                }
            }
            sr.Close();
        }

        private void LoadClassTemplate()
        {
            classTemplate = string.Empty;

            // ���ø� ������ �ִ��� üũ�Ѵ�.
            if (!File.Exists(PATH_CLASS_TEMPLATE))
            {
                MessageBox.Show("Ŭ���� ���ø� ������ ã�� �� �����ϴ�!");
                return;
            }

            // ���ø� ������ ������ �ε��ؿ´�.
            var sr = new StreamReader(PATH_CLASS_TEMPLATE);
            classTemplate = sr.ReadToEnd();
            sr.Close();
        }
        #endregion

        private void Btn_LoadExcelListPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("���� ����Ʈ �ҷ�����", "TB_ExcelList", false, "txt");
        }

        private void Btn_LoadExcelPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("���� ���� ��ġ �ҷ�����", "TB_ExcelPath");
        }

        private void Btn_SaveCsPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("CS ���� ���� ��� ����", "TB_CsPath");
        }

        private void Btn_SaveJsonPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("Json ���� ���� ��� ����", "TB_JsonPath");
        }

        private void Btn_ConvertToCs_Click(object sender, EventArgs e)
        {
            ReadExcelFiles();
            ConvertToCs();
            EndConvert();
        }

        private void Btn_ConvertToJson_Click(object sender, EventArgs e)
        {
            ReadExcelFiles();
            ConvertToJson();
            EndConvert();
        }

        private void Btn_ConvertAll_Click(object sender, EventArgs e)
        {
            ReadExcelFiles();
            ConvertAll();
            EndConvert();
        }

        private void OpenFileDialog(
            string title, string textBoxName, bool isFolder = true, string? extension = null)
        {
            if (!fieldByName.ContainsKey(textBoxName))
            {
                MessageBox.Show($"��ȿ���� ���� �ؽ�Ʈ �ڽ� ID �Դϴ�. => {textBoxName}");
                return;
            }

            var dialog = new CommonOpenFileDialog(title);
            dialog.IsFolderPicker = isFolder;
            if (!isFolder)
            {
                dialog.DefaultFileName = $"*.{extension}";

                var filter = new CommonFileDialogFilter();
                filter.DisplayName = extension;
                filter.Extensions.Add(extension);
                dialog.Filters.Add(filter);
            }

            var result = dialog.ShowDialog();
            switch (result)
            {
                case CommonFileDialogResult.Ok:
                    var textBox = fieldByName[textBoxName].GetValue(this) as TextBox;
                    if (textBox == null)
                    {
                        MessageBox.Show($"{textBoxName} �ʵ� ����� �������µ� �����߽��ϴ�.");
                        return;
                    }
                    textBox.Text = dialog.FileName;
                    SavePath(textBoxName, textBox.Text);
                    break;
                default:
                    break;
            }
            dialog.Dispose();
        }

        private void SavePath(string textBoxName, string path)
        {
            var pathByKey = new Dictionary<string, string>();

            if (File.Exists(PATH_SAVED))
            {
                // ������ �����Ѵٸ� �о�� dict�� �����Ѵ�.
                var sr = new StreamReader(PATH_SAVED);
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (line == null) continue;
                    var words = line.Split('=');
                    pathByKey.Add(words[0], words[1]);
                }
                sr.Close();
            }
            else
            {
                // ������ �������� �ʴ´ٸ� �������ش�.
                File.Create(PATH_SAVED);
            }

            // �� ��� ������ �ۼ��Ѵ�.
            if (pathByKey.ContainsKey(textBoxName))
                pathByKey[textBoxName] = path;
            else
                pathByKey.Add(textBoxName, path);

            // Dictionary�� �ؽ�Ʈ�� ��ȯ�� ���Ϸ� �����Ѵ�.
            var sw = new StreamWriter(PATH_SAVED);
            foreach(var line in pathByKey)
                sw.WriteLine($"{line.Key}={line.Value}");
            sw.Close();
        }

        private void ReadExcelFiles()
        {
            // �о�� ���� ���� ����Ʈ�� �ִ��� üũ�Ѵ�.
            var excelListPath = TB_ExcelList.Text;
            if (!File.Exists(excelListPath))
            {
                MessageBox.Show($"���� ���� ����Ʈ�� ��ΰ� �������� �ʾҰų� ��ȿ���� �ʽ��ϴ�!");
                return;
            }

            // ���� ���� ����Ʈ�� �ۼ��Ѵ�.
            UpdateDescText("��ȯ�� ���� ����Ʈ �о������...");
            var fileNameList = new List<string>();
            var sr = new StreamReader(excelListPath);
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (line == null) continue;
                fileNameList.Add(line);
            }
            sr.Close();

            // ����Ʈ�� ���� ������ DataTableCollection���� �����Ѵ�.
            int progress = 0;
            UpdateDescText($"���� �д���...({progress}/{fileNameList.Count})");
            var threadList = new List<Thread>();
            for(int i = 0; i < fileNameList.Count; i++)
            {
                var thread = new Thread((idx) =>
                {
                    int iFile = (int)idx;

                    // ���� �� ���� ��θ� ���´�.
                    var fileName = fileNameList[iFile];
                    var dirPath = TB_ExcelPath.Text;

                    // �ش��ϴ� ���� ������ �����ϴ��� üũ�Ѵ�.
                    var path = Path.Combine(dirPath, fileName);
                    if (!File.Exists(path))
                    {
                        MessageBox.Show($"{path}: �ش��ϴ� �̸��� ���� ������ �����ϴ�!");
                        return;
                    }

                    // ���� ������ �о� DataTableCollection���� ��ȯ�� �����Ѵ�.
                    DataTableCollection? excel = null;
                    using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet();
                            excel = result.Tables;
                        }
                    }

                    lock (lockObj)
                    {
                        // ���̺� ������ �߰��Ѵ�.
                        excelList.Add(excel);

                        // �˸� �ؽ�Ʈ�� �����Ѵ�.
                        progress++;
                        UpdateDescText($"���� �д���...({progress}/{fileNameList.Count})");
                    }
                });
                thread.IsBackground = true;
                thread.Start(i);
                threadList.Add(thread);
            }
            foreach(var thread in threadList) thread.Join();
        }

        private void ConvertToCs()
        {
            // ������ ��ġ�� ��ȿ���� üũ�Ѵ�.
            if (!Directory.Exists(TB_CsPath.Text))
            {
                MessageBox.Show("cs ������ ������ ���丮 ��ΰ� ��ȿ���� �ʽ��ϴ�!");
                return;
            }

            // ������ ��ȸ�Ѵ�.
            var threadList = new List<Thread>();
            int progress = 0;
            foreach (var excel in excelList)
            {
                var thread = new Thread((_excel) =>
                {
                    var tableCollection = _excel as DataTableCollection;
                    if (tableCollection == null) return;

                    foreach (DataTable table in tableCollection)
                    {
                        string tableName = table.TableName;
                        string rowName = $"{tableName}Row";

                        // ���̺��� ���� cs���� ������ �������ش�.
                        var contentSB = new StringBuilder(classTemplate);
                        contentSB.Replace("#TableName#", tableName);
                        contentSB.Replace("#TableRow#", rowName);
                        contentSB.Append("\r\n\r\n");

                        // Row Ŭ���� ������ �ۼ��Ѵ�.
                        contentSB.Append("[Serializable]\r\n");
                        contentSB.Append($"public class {rowName} : ITableRow\r\n");
                        contentSB.Append("{\r\n");
                        var keys = table.Rows[0];
                        var types = table.Rows[1];
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            // ���������� üũ�� ������ ��쿡 ���� Ŭ������ �ۼ��Ѵ�.
                            var key = keys[i].ToString();
                            var type = types[i].ToString();
                            var splitted = type.Split('_');
                            if (splitted[0].Equals("#E"))
                                contentSB.Append($"\t[JsonProperty, JsonConverter(typeof(StringEnumConverter))] public readonly {splitted[1]} {key};\r\n");
                            else
                                contentSB.Append($"\t[JsonProperty] public readonly {type} {key};\r\n");
                        }
                        contentSB.Append("\r\n");
                        contentSB.Append($"\tpublic string GetKey() {{ return {keys[0]}.ToString(); }}\r\n");
                        contentSB.Append("}\r\n");

                        // ���������� �ۼ��� ������ Cs ���Ϸ� �����Ѵ�.
                        string saveDir = TB_CsPath.Text;
                        string path = Path.Combine(saveDir, $"{tableName}.cs");
                        File.Create(path).Close();
                        var sw = new StreamWriter(path);
                        sw.Write(contentSB.ToString());
                        sw.Close();

                        // �ȳ� �ؽ�Ʈ�� �����Ѵ�.
                        lock (lockObj)
                        {
                            progress++;
                            UpdateDescText($"CS ���Ϸ� ��ȯ��...({progress}/{excelList.Count})");
                        }
                    }
                });
                thread.IsBackground = true;
                thread.Start(excel);
                threadList.Add(thread);
            }
            foreach (var thread in threadList) thread.Join();
        }

        private void ConvertToJson()
        {
            // ������ ��ġ�� ��ȿ���� üũ�Ѵ�.
            if (!Directory.Exists(TB_JsonPath.Text))
            {
                MessageBox.Show("json ������ ������ ���丮 ��ΰ� ��ȿ���� �ʽ��ϴ�!");
                return;
            }

            // ������ ��ȸ�Ѵ�.
            var threadList = new List<Thread>();
            int progress = 0;
            foreach (var excel in excelList)
            {
                var thread = new Thread((_excel) =>
                {
                    var tableCollection = _excel as DataTableCollection;
                    if (tableCollection == null) return;

                    foreach (DataTable table in tableCollection)
                    {
                        string tableName = table.TableName;

                        // Json ������ �ۼ��Ѵ�.
                        var json = new JObject();
                        var jArr = new JArray();
                        var keys = table.Rows[0];
                        var types = table.Rows[1];
                        for (int i = 2; i < table.Rows.Count; i++)
                        {
                            var jObj = new JObject();
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                var key = keys[j].ToString();
                                var typeInfo = types[j].ToString();
                                var type = typeInfo.Split('_')[0];
                                switch (type)
                                {
                                    case "BigInteger":
                                    case "string":
                                    case "#E":
                                        jObj.Add(key, table.Rows[i][j].ToString());
                                        break;
                                    case "float":
                                    case "double":
                                        jObj.Add(key, float.Parse(table.Rows[i][j].ToString()));
                                        break;
                                    default:
                                        jObj.Add(key, int.Parse(table.Rows[i][j].ToString()));
                                        break;
                                }
                            }
                            jArr.Add(jObj);
                        }
                        json.Add("rows", jArr);

                        // ���������� �ۼ��� ������ Json ���Ϸ� �����Ѵ�.
                        string saveDir = TB_JsonPath.Text;
                        string path = Path.Combine(saveDir, $"{tableName}.json");
                        File.Create(path).Close();
                        var sw = new StreamWriter(path);
                        sw.Write(json.ToString());
                        sw.Close();

                        // �ȳ� �ؽ�Ʈ�� �����Ѵ�.
                        lock (lockObj)
                        {
                            progress++;
                            UpdateDescText($"CS ���Ϸ� ��ȯ��...({progress}/{excelList.Count})");
                        }
                    }
                });
                thread.IsBackground = true;
                thread.Start(excel);
                threadList.Add(thread);
            }
            foreach (var thread in threadList) thread.Join();
        }

        private void ConvertAll()
        {
            // ������ ��ġ�� ��ȿ���� üũ�Ѵ�.
            if (!Directory.Exists(TB_CsPath.Text))
            {
                MessageBox.Show("������ ������ ���丮 ��ΰ� ��ȿ���� �ʽ��ϴ�!");
                return;
            }

            if (!Directory.Exists(TB_JsonPath.Text))
            {
                MessageBox.Show("json ������ ������ ���丮 ��ΰ� ��ȿ���� �ʽ��ϴ�!");
                return;
            }

            // ������ ��ȸ�Ѵ�.
            var threadList = new List<Thread>();
            int progress = 0;
            foreach (var excel in excelList)
            {
                var thread = new Thread((_excel) =>
                {
                    var tableCollection = _excel as DataTableCollection;
                    if (tableCollection == null) return;

                    foreach (DataTable table in tableCollection)
                    {
                        string tableName = table.TableName;
                        string rowName = $"{tableName}Row";

                        // ���̺��� ���� cs���� ������ �������ش�.
                        var contentSB = new StringBuilder(classTemplate);
                        contentSB.Replace("#TableName#", tableName);
                        contentSB.Replace("#TableRow#", rowName);
                        contentSB.Append("\r\n\r\n");

                        // Row Ŭ���� ������ �ۼ��Ѵ�.
                        contentSB.Append("[Serializable]\r\n");
                        contentSB.Append($"public class {rowName} : ITableRow\r\n");
                        contentSB.Append("{\r\n");
                        var keys = table.Rows[0];
                        var types = table.Rows[1];
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            // ���������� üũ�� ������ ��쿡 ���� Ŭ������ �ۼ��Ѵ�.
                            var key = keys[i].ToString();
                            var type = types[i].ToString();
                            var splitted = type.Split('_');
                            if (splitted[0].Equals("#E"))
                                contentSB.Append($"\t[JsonProperty, JsonConverter(typeof(StringEnumConverter))] public readonly {splitted[1]} {key};\r\n");
                            else
                                contentSB.Append($"\t[JsonProperty] public readonly {type} {key};\r\n");
                        }
                        contentSB.Append("\r\n");
                        contentSB.Append($"\tpublic string GetKey() {{ return {keys[0]}.ToString(); }}\r\n");
                        contentSB.Append("}\r\n");

                        // ���������� �ۼ��� ������ Cs ���Ϸ� �����Ѵ�.
                        string saveDir = TB_CsPath.Text;
                        string path = Path.Combine(saveDir, $"{tableName}.cs");
                        File.Create(path).Close();
                        var sw = new StreamWriter(path);
                        sw.Write(contentSB.ToString());
                        sw.Close();

                        // Json ������ �ۼ��Ѵ�.
                        var json = new JObject();
                        var jArr = new JArray();
                        for (int i = 2; i < table.Rows.Count; i++)
                        {
                            var jObj = new JObject();
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                var key = keys[j].ToString();
                                var typeInfo = types[j].ToString();
                                var type = typeInfo.Split('_')[0];
                                switch (type)
                                {
                                    case "BigInteger":
                                    case "string":
                                    case "#E":
                                        jObj.Add(key, table.Rows[i][j].ToString());
                                        break;
                                    case "float":
                                    case "double":
                                        jObj.Add(key, float.Parse(table.Rows[i][j].ToString()));
                                        break;
                                    default:
                                        jObj.Add(key, int.Parse(table.Rows[i][j].ToString()));
                                        break;
                                }
                            }
                            jArr.Add(jObj);
                        }
                        json.Add("rows", jArr);

                        // ���������� �ۼ��� ������ Json ���Ϸ� �����Ѵ�.
                        saveDir = TB_JsonPath.Text;
                        path = Path.Combine(saveDir, $"{tableName}.json");
                        File.Create(path).Close();
                        sw = new StreamWriter(path);
                        sw.Write(json.ToString());
                        sw.Close();

                        // �ȳ� �ؽ�Ʈ�� �����Ѵ�.
                        lock (lockObj)
                        {
                            progress++;
                            UpdateDescText($"Cs/Json ���Ϸ� ��ȯ��...({progress}/{excelList.Count})");
                        }
                    }
                });
                thread.IsBackground = true;
                thread.Start(excel);
                threadList.Add(thread);
            }
            foreach (var thread in threadList) thread.Join();
        }

        private void EndConvert()
        {
            excelList.Clear();
            UpdateDescText("��ȯ �Ϸ�!");
        }

        private void UpdateDescText(string desc)
        {
            Label_Desc.Text = desc;
        }
    }
}
