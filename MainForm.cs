using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WordHelperLibrary;

namespace WordMerger
{
    public partial class MainForm : Form
    {
        #region constructor
        public MainForm()
        {
            InitializeComponent();

            InitUi();
        }
        #endregion

        #region property
        private TextBox _txtFileList;
        private TextBox _txtLog;

        private const int ControlMargin = 20;
        private const int ControlPadding = 12;
        #endregion

        #region event handler
        private void BtnAddFile_Click(object sender, EventArgs e)
        {
            var openDlg = new OpenFileDialog { Filter = "Word文件(*.docx)|*.docx|所有文件(*.*)|*.*", Multiselect = true };
            if (openDlg.ShowDialog() != DialogResult.OK) return;
            var inputFileList = openDlg.FileNames.ToList();
            foreach (var fileName in inputFileList)
            {
                _txtFileList.AppendText($"{fileName}\r\n");
            }
        }

        private void BtnMerge_Click(object sender, EventArgs e)
        {
            var inputFilenameList = _txtFileList.Lines.ToList();
            inputFilenameList.RemoveAll(string.IsNullOrWhiteSpace);
            if (inputFilenameList.Count == 0)
            {
                _txtLog.AppendText("未添加需要合并的Word文件\r\n");
                return;
            }
            var path = Path.GetDirectoryName(inputFilenameList.First());
            var ext = Path.GetExtension(inputFilenameList.First());
            var outputFileName = Path.Combine(path, $"MergedFile - {DateTime.Now:yyyyMMddHHmmssfff}{ext}");
            var s = MergeHelper.Merge(inputFilenameList, outputFileName);
            if (string.IsNullOrWhiteSpace(s)) _txtLog.AppendText("合并完成\r\n");
            else _txtLog.AppendText($"{s}\r\n");
        }

        private void txtFileList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void txtFileList_DragDrop(object sender, DragEventArgs e)
        {
            string[] fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            foreach (var fileName in fileList)
            {
                _txtFileList.AppendText($"{fileName}\r\n");
            }
        }

        #endregion

        #region ui
        private void InitUi()
        {
            MaximizeBox = false;
            ShowIcon = true;
            StartPosition = FormStartPosition.CenterScreen;
            Text = $"Word文档合并器 {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}";

            var btnAddFile = new Button
            {
                AutoSize = true,
                Location = new Point(ControlMargin, ControlMargin),
                Parent = this,
                Text = "添加文件"
            };
            btnAddFile.Click += BtnAddFile_Click;

            var btnMerge = new Button
            {
                AutoSize = true,
                Location = new Point(btnAddFile.Right + ControlPadding, btnAddFile.Top),
                Parent = this,
                Text = "开始合并"
            };
            btnMerge.Click += BtnMerge_Click;

            _txtLog = new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right,
                Location = new Point(btnAddFile.Left, ClientSize.Height - ControlMargin - 100),
                Multiline = true,
                Parent = this,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Size = new Size(ClientSize.Width - 2 * btnAddFile.Left, 100),
                WordWrap = false
            };

            _txtFileList = new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                Location = new Point(btnAddFile.Left, btnAddFile.Bottom + ControlPadding),
                Multiline = true,
                Parent = this,
                ScrollBars = ScrollBars.Vertical,
                Size = new Size(ClientSize.Width - 2 * btnAddFile.Left, _txtLog.Top - btnAddFile.Bottom - 2 * ControlPadding),
                WordWrap = false
            };
            _txtFileList.AllowDrop = true;
            _txtFileList.DragDrop += txtFileList_DragDrop;
            _txtFileList.DragEnter += txtFileList_DragEnter;
        }
        #endregion
    }
}
