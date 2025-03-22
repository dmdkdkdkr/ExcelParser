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
            // .Net 5.0 이상부터 생기는 ExcelDataReader의 인코딩 버그 수정용
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void CacheFields()
        {
            var fields = typeof(Form_ExcelParser).GetFields(BindingFlags.Instance | BindingFlags.NonPublic);
            foreach (var f in fields) fieldByName.Add(f.Name, f);
        }

        private void LoadSavedPaths()
        {
            // 파일의 존재 여부를 체크한다. 없다면 생성후 리턴한다.
            if (!File.Exists(PATH_SAVED))
            {
                File.Create(PATH_SAVED).Dispose();
                return;
            }

            // 파일이 존재한다면 읽어와 dict로 저장한다.
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

            // 템플릿 파일이 있는지 체크한다.
            if (!File.Exists(PATH_CLASS_TEMPLATE))
            {
                MessageBox.Show("클래스 템플릿 파일을 찾을 수 없습니다!");
                return;
            }

            // 템플릿 파일의 내용을 로드해온다.
            var sr = new StreamReader(PATH_CLASS_TEMPLATE);
            classTemplate = sr.ReadToEnd();
            sr.Close();
        }
        #endregion

        private void Btn_LoadExcelListPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("엑셀 리스트 불러오기", "TB_ExcelList", false, "txt");
        }

        private void Btn_LoadExcelPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("엑셀 파일 위치 불러오기", "TB_ExcelPath");
        }

        private void Btn_SaveCsPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("CS 파일 저장 경로 지정", "TB_CsPath");
        }

        private void Btn_SaveJsonPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog("Json 파일 저장 경로 지정", "TB_JsonPath");
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
                MessageBox.Show($"유효하지 않은 텍스트 박스 ID 입니다. => {textBoxName}");
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
                        MessageBox.Show($"{textBoxName} 필드 멤버를 가져오는데 실패했습니다.");
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
                // 파일이 존재한다면 읽어와 dict로 저장한다.
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
                // 파일이 존재하지 않는다면 생성해준다.
                File.Create(PATH_SAVED);
            }

            // 새 경로 정보를 작성한다.
            if (pathByKey.ContainsKey(textBoxName))
                pathByKey[textBoxName] = path;
            else
                pathByKey.Add(textBoxName, path);

            // Dictionary를 텍스트로 변환후 파일로 저장한다.
            var sw = new StreamWriter(PATH_SAVED);
            foreach(var line in pathByKey)
                sw.WriteLine($"{line.Key}={line.Value}");
            sw.Close();
        }

        private void ReadExcelFiles()
        {
            // 읽어올 엑셀 파일 리스트가 있는지 체크한다.
            var excelListPath = TB_ExcelList.Text;
            if (!File.Exists(excelListPath))
            {
                MessageBox.Show($"엑셀 파일 리스트의 경로가 지정되지 않았거나 유효하지 않습니다!");
                return;
            }

            // 엑셀 파일 리스트를 작성한다.
            UpdateDescText("변환할 엑셀 리스트 읽어오는중...");
            var fileNameList = new List<string>();
            var sr = new StreamReader(excelListPath);
            while (!sr.EndOfStream)
            {
                var line = sr.ReadLine();
                if (line == null) continue;
                fileNameList.Add(line);
            }
            sr.Close();

            // 리스트를 따라 엑셀을 DataTableCollection으로 저장한다.
            int progress = 0;
            UpdateDescText($"엑셀 읽는중...({progress}/{fileNameList.Count})");
            var threadList = new List<Thread>();
            for(int i = 0; i < fileNameList.Count; i++)
            {
                var thread = new Thread((idx) =>
                {
                    int iFile = (int)idx;

                    // 폴더 및 파일 경로를 얻어온다.
                    var fileName = fileNameList[iFile];
                    var dirPath = TB_ExcelPath.Text;

                    // 해당하는 엑셀 파일이 존재하는지 체크한다.
                    var path = Path.Combine(dirPath, fileName);
                    if (!File.Exists(path))
                    {
                        MessageBox.Show($"{path}: 해당하는 이름의 엑셀 파일이 없습니다!");
                        return;
                    }

                    // 엑셀 파일을 읽어 DataTableCollection으로 변환해 저장한다.
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
                        // 테이블 정보를 추가한다.
                        excelList.Add(excel);

                        // 알림 텍스트를 갱신한다.
                        progress++;
                        UpdateDescText($"엑셀 읽는중...({progress}/{fileNameList.Count})");
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
            // 저장할 위치가 유효한지 체크한다.
            if (!Directory.Exists(TB_CsPath.Text))
            {
                MessageBox.Show("cs 파일을 저장할 디렉토리 경로가 유효하지 않습니다!");
                return;
            }

            // 엑셀을 순회한다.
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

                        // 테이블을 따라 cs파일 내용을 설정해준다.
                        var contentSB = new StringBuilder(classTemplate);
                        contentSB.Replace("#TableName#", tableName);
                        contentSB.Replace("#TableRow#", rowName);
                        contentSB.Append("\r\n\r\n");

                        // Row 클래스 내용을 작성한다.
                        contentSB.Append("[Serializable]\r\n");
                        contentSB.Append($"public class {rowName} : ITableRow\r\n");
                        contentSB.Append("{\r\n");
                        var keys = table.Rows[0];
                        var types = table.Rows[1];
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            // 열거형인지 체크후 각각의 경우에 따라 클래스를 작성한다.
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

                        // 최종적으로 작성된 내용을 Cs 파일로 저장한다.
                        string saveDir = TB_CsPath.Text;
                        string path = Path.Combine(saveDir, $"{tableName}.cs");
                        File.Create(path).Close();
                        var sw = new StreamWriter(path);
                        sw.Write(contentSB.ToString());
                        sw.Close();

                        // 안내 텍스트를 설정한다.
                        lock (lockObj)
                        {
                            progress++;
                            UpdateDescText($"CS 파일로 변환중...({progress}/{excelList.Count})");
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
            // 저장할 위치가 유효한지 체크한다.
            if (!Directory.Exists(TB_JsonPath.Text))
            {
                MessageBox.Show("json 파일을 저장할 디렉토리 경로가 유효하지 않습니다!");
                return;
            }

            // 엑셀을 순회한다.
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

                        // Json 내용을 작성한다.
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

                        // 최종적으로 작성된 내용을 Json 파일로 저장한다.
                        string saveDir = TB_JsonPath.Text;
                        string path = Path.Combine(saveDir, $"{tableName}.json");
                        File.Create(path).Close();
                        var sw = new StreamWriter(path);
                        sw.Write(json.ToString());
                        sw.Close();

                        // 안내 텍스트를 설정한다.
                        lock (lockObj)
                        {
                            progress++;
                            UpdateDescText($"CS 파일로 변환중...({progress}/{excelList.Count})");
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
            // 저장할 위치가 유효한지 체크한다.
            if (!Directory.Exists(TB_CsPath.Text))
            {
                MessageBox.Show("파일을 저장할 디렉토리 경로가 유효하지 않습니다!");
                return;
            }

            if (!Directory.Exists(TB_JsonPath.Text))
            {
                MessageBox.Show("json 파일을 저장할 디렉토리 경로가 유효하지 않습니다!");
                return;
            }

            // 엑셀을 순회한다.
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

                        // 테이블을 따라 cs파일 내용을 설정해준다.
                        var contentSB = new StringBuilder(classTemplate);
                        contentSB.Replace("#TableName#", tableName);
                        contentSB.Replace("#TableRow#", rowName);
                        contentSB.Append("\r\n\r\n");

                        // Row 클래스 내용을 작성한다.
                        contentSB.Append("[Serializable]\r\n");
                        contentSB.Append($"public class {rowName} : ITableRow\r\n");
                        contentSB.Append("{\r\n");
                        var keys = table.Rows[0];
                        var types = table.Rows[1];
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            // 열거형인지 체크후 각각의 경우에 따라 클래스를 작성한다.
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

                        // 최종적으로 작성된 내용을 Cs 파일로 저장한다.
                        string saveDir = TB_CsPath.Text;
                        string path = Path.Combine(saveDir, $"{tableName}.cs");
                        File.Create(path).Close();
                        var sw = new StreamWriter(path);
                        sw.Write(contentSB.ToString());
                        sw.Close();

                        // Json 내용을 작성한다.
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

                        // 최종적으로 작성된 내용을 Json 파일로 저장한다.
                        saveDir = TB_JsonPath.Text;
                        path = Path.Combine(saveDir, $"{tableName}.json");
                        File.Create(path).Close();
                        sw = new StreamWriter(path);
                        sw.Write(json.ToString());
                        sw.Close();

                        // 안내 텍스트를 설정한다.
                        lock (lockObj)
                        {
                            progress++;
                            UpdateDescText($"Cs/Json 파일로 변환중...({progress}/{excelList.Count})");
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
            UpdateDescText("변환 완료!");
        }

        private void UpdateDescText(string desc)
        {
            Label_Desc.Text = desc;
        }
    }
}
