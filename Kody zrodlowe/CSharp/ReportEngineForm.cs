using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ReportingSystem.Engine
{
    /// <summary>
    /// Dynamic Reporting System Engine
    /// Generates UI forms and SQL queries based on external report definition files (*.dr)
    /// </summary>
    public partial class ReportEngineForm : Form
    {
        enum DataType { dtNone = 0, dtString = 1, dtInteger = 2, dtFloat = 3, dtDate = 4, dtDateTime = 5, dtSelect = 9 };

        private dbQuery _reportQuery;
        private const string HelpFile = "Help.mht";

        public delegate void RecordDoubleClickedHandler(string reportName, DataRow row);
        public event RecordDoubleClickedHandler OnRecordDoubleClicked = null;

        private class ReportParameter
        {
            public int Index;
            public DataType ParamType;
            public string Name;
            public string Description;
            public bool IsRequired;
            public string DefaultValue;
            public string CalculatedSqlValue;

            // Dynamic UI Components
            public Panel ContainerPanel;
            public Label TitleLabel;
            public CheckBox OptionalCheckBox;
            public Label OperatorLabel;
            public TextBox TextControl;
            public TextBox IntegerControl;
            public TextBox FloatControl;
            public DateTimePicker DateControl;
            public DateTimePicker DateTimeControl;
            public ComboBox SelectControl;

            public int DictionaryIndex;
        }

        private class ReportDictionary
        {
            public int Index;
            public string Name;
            public string SqlBlockName;
            public string KeyValue;
            public string FieldKey;
            public string FieldDescription;
            public string FieldForeign;
            public string SelectSql;
            public DataTable Data;
        }

        // Configuration and Styling
        private string _reportsDirectory = null;
        private Color _colorOptionalParam = SystemColors.GradientActiveCaption;
        private Color _colorRequiredParam = SystemColors.Info;
        private Color _gridRowColorA = SystemColors.GradientActiveCaption;
        private Color _gridRowColorB = SystemColors.GradientInactiveCaption;

        private string[] _availableReportFiles = null;
        private List<ReportParameter> _parameters = new List<ReportParameter>();
        private List<ReportDictionary> _dictionaries = new List<ReportDictionary>();

        private string _currentReportPath = null;
        private string _rawReportContent = null;
        private string _sqlBlock = string.Empty;
        private DataTable _reportResults = new DataTable();

        public ReportEngineForm()
        {
            InitializeComponent();
            SetupDefaultEnvironment();
        }

        private void SetupDefaultEnvironment()
        {
            _reportsDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reports");

            dgvData.ReadOnly = true;
            dgvData.AutoGenerateColumns = true;

            // Initialize UI tabs state
            tabData.TabPages.Remove(tabPageData);
            tabData.TabPages.Remove(tabPageSQL);

            UpdateControlStates();
        }

        #region --- Report Processing Logic ---

        private bool LoadReport(string filePath)
        {
            try
            {
                panelBody.Visible = false;

                // Load file content with encoding detection (simplified for presentation)
                _rawReportContent = File.ReadAllText(filePath, Encoding.UTF8);

                // Parse core blocks
                string paramBlock = ExtractBlock("<PARAM>", "</PARAM>", _rawReportContent);
                string[] paramLines = paramBlock.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);

                _sqlBlock = PrepareSqlParams(ExtractBlock("<SQL>", "</SQL>", _rawReportContent));

                string description = ExtractBlock("<DESCRIPTION>", "</DESCRIPTION>", _rawReportContent);
                lblReportDescription.Text = description;

                // Parse Dictionaries
                ParseDictionaries(_rawReportContent);

                // Parse Parameters
                foreach (string line in paramLines)
                {
                    ParseParameterLine(line);
                }

                RenderDynamicParameters();
                InitializeSelectControls();

                _currentReportPath = filePath;
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"Error loading report: {ex.Message}", Color.Red);
                return false;
            }
            finally
            {
                panelBody.Visible = true;
            }
        }

        /// <summary>
        /// Logic for processing <IFPARAM> and <IFWHERE> tags in SQL
        /// </summary>
        private string BuildFinalSqlQuery(string template)
        {
            string query = template;

            // Process conditional WHERE blocks
            string whereBlock;
            do
            {
                whereBlock = ExtractBlock("<IFWHERE>", "</IFWHERE>", query, false);
                if (!string.IsNullOrEmpty(whereBlock))
                    query = query.Replace(whereBlock, ProcessSqlConditionalLogic(whereBlock));
            } while (!string.IsNullOrEmpty(whereBlock));

            // Inject parameter values
            foreach (var param in _parameters)
            {
                bool isActive = (param.OptionalCheckBox == null) || (param.OptionalCheckBox.Checked);
                if (isActive)
                {
                    string placeholder = $":{param.Name}";
                    query = query.Replace(placeholder, param.CalculatedSqlValue);
                }
            }

            return query;
        }

        private string ProcessSqlConditionalLogic(string source)
        {
            // Logic for handling <IFPARAM> tags based on user selection
            // [Included for context: handles whether to include SQL fragments based on provided params]
            return source; // Simplified for the snippet
        }

        #endregion

        #region --- UI Rendering ---

        private void RenderDynamicParameters()
        {
            // Reverse order rendering for Top-Docking
            for (int i = _parameters.Count - 1; i >= 0; i--)
            {
                var param = _parameters[i];
                param.ContainerPanel = new Panel
                {
                    Parent = this.panelBody,
                    Dock = DockStyle.Top,
                    Height = 30,
                    BorderStyle = BorderStyle.Fixed3D,
                    Tag = i
                };

                // Create Label
                param.TitleLabel = new Label
                {
                    Parent = param.ContainerPanel,
                    Text = param.Description,
                    BackColor = param.IsRequired ? _colorRequiredParam : _colorOptionalParam,
                    Width = 150,
                    TextAlign = ContentAlignment.MiddleLeft
                };

                // Logic for creating TextBoxes/DateTimePickers based on DataType...
                // (Matches the logic in your original file)
            }
        }

        #endregion

        #region --- Events & Actions ---

        private void btnRunReport_Click(object sender, EventArgs e)
        {
            if (ValidateParameters())
            {
                try
                {
                    string finalSql = BuildFinalSqlQuery(_sqlBlock);
                    txtSqlPreview.Text = finalSql;

                    // Execution of the report
                    ExecuteReportQuery(finalSql);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Execution error: {ex.Message}");
                }
            }
        }

        private void ExecuteReportQuery(string sql)
        {
            // Placeholder for Database Manager execution logic
            // After execution:
            // dgvData.DataSource = results;
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            // Integration with ExportToExcel library
            LogMessage("Exporting to Excel...", Color.Blue);
        }

        #endregion

        #region --- Utility Methods ---

        private string ExtractBlock(string startTag, string endTag, string source, bool removeTags = true)
        {
            try
            {
                // Case-insensitive tag extraction
                int p1 = source.IndexOf(startTag, StringComparison.OrdinalIgnoreCase);
                int p2 = source.IndexOf(endTag, StringComparison.OrdinalIgnoreCase);

                if (p1 != -1 && p2 > p1)
                {
                    int contentStart = p1 + startTag.Length;
                    if (removeTags)
                        return source.Substring(contentStart, p2 - contentStart).Trim();
                    else
                        return source.Substring(p1, (p2 + endTag.Length) - p1);
                }
            }
            catch { }
            return string.Empty;
        }

        private void LogMessage(string message, Color color)
        {
            // RichTextBox logging logic
            rtbMessages.SelectionColor = color;
            rtbMessages.AppendText($"{DateTime.Now:HH:mm:ss} - {message}\n");
        }

        private void UpdateControlStates()
        {
            bool hasReportLoaded = (_rawReportContent != null);
            btnRunReport.Enabled = hasReportLoaded;
            btnCloseReport.Enabled = hasReportLoaded;
        }

        #endregion
    }

    /// <summary>
    /// Helper class for handling ComboBox items with Key-Value pairs
    /// </summary>
    public class DictionaryItem
    {
        public string Text { get; set; }
        public object Value { get; set; }
        public override string ToString() => Text;
    }
}