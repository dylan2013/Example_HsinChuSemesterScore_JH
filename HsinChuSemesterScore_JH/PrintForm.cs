using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using JHSchool.Data;
using System.IO;
using Aspose.Words;
using JHSchool.Evaluation.Calculation;
using JHSchool.Evaluation.Mapping;

namespace HsinChuSemesterScore_JH
{
    public partial class PrintForm : BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        // 學期科目名稱
        private List<string> _SemesterSubjectNameList = new List<string>();
        // 畫面上所選的學生編號
        List<string> _StudentIDList;
        // 缺曠區間統計
        Dictionary<string, Dictionary<string, int>> _AttendanceDict = new Dictionary<string, Dictionary<string, int>>();
        // 學生日常生活表現XML
        Dictionary<string, DAO.StudTextScoreXML> _StudTextScoreXMLDict = new Dictionary<string, DAO.StudTextScoreXML>();

        // 學生學期成績
        Dictionary<string, DAO.StudentScore> _StudentScoreDict = new Dictionary<string, DAO.StudentScore>();

        private List<string> typeList = new List<string>();
        private List<string> absenceList = new List<string>();
        private List<string> _SelSubjNameList = new List<string>();
        private List<string> _SelAttendanceList = new List<string>();
        
        // 產生報表用
        private BackgroundWorker _bgWorkReport;
        private DocumentBuilder _builder;

        // 載入讀取資料用
        BackgroundWorker bkw;

        // 錯誤訊息
        List<string> _ErrorList = new List<string>();

        // 領域錯誤訊息
        List<string> _ErrorDomainNameList = new List<string>();

        // 樣板內有科目名稱
        List<string> _TemplateSubjectNameList = new List<string>();

        // 存檔路徑
        string pathW = "";

        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();

        // 開始日期
        private DateTime _BeginDate;
        // 結束日期
        private DateTime _EndDate;

        // 成績校正日期字串
        private string _ScoreEditDate = "";

        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        // 畫面上選的學年度
        private int _SelSchoolYear;
        // 畫面上選的學期
        private int _SelSemester;
        // 不排名
        private string _SelNotRankedFilter = "";

        // 等第對照
        private DegreeMapper _degreeMapper;

        private Dictionary<string, List<string>> _StudTagDict = new Dictionary<string, List<string>>();

        // 紀錄樣板設定
        List<DAO.UDT_ScoreConfig> _UDTConfigList;

        public PrintForm(List<string> StudIDList)
        {
            InitializeComponent();            
            _StudentIDList = StudIDList;
            _degreeMapper = new DegreeMapper();
            bkw = new BackgroundWorker();
            bkw.DoWork += new DoWorkEventHandler(bkw_DoWork);
            bkw.ProgressChanged += new ProgressChangedEventHandler(bkw_ProgressChanged);
            bkw.WorkerReportsProgress = true;
            bkw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bkw_RunWorkerCompleted);
            _bgWorkReport = new BackgroundWorker();
            _bgWorkReport.DoWork += new DoWorkEventHandler(_bgWorkReport_DoWork);
            _bgWorkReport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorkReport_RunWorkerCompleted);
            _bgWorkReport.WorkerReportsProgress = true;
            _bgWorkReport.ProgressChanged += new ProgressChangedEventHandler(_bgWorkReport_ProgressChanged);
        }

        void bkw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            circularProgress1.Value = e.ProgressPercentage;
        }

        void bkw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            EnbSelect();

            _DefalutSchoolYear = K12.Data.School.DefaultSchoolYear;
            _DefaultSemester = K12.Data.School.DefaultSemester;

            if (_Configure == null)
                _Configure = new Configure();

            cboConfigure.Items.Clear();
            foreach (var item in _ConfigureList)
            {
                cboConfigure.Items.Add(item);
            }
            cboConfigure.Items.Add(new Configure() { Name = "新增" });
            int i;

            if (int.TryParse(_DefalutSchoolYear, out i))
            {
                for (int j = 5; j > 0; j--)
                {
                    cboSchoolYear.Items.Add("" + (i - j));
                }
                
                for (int j = 0; j < 3; j++)
                {
                    cboSchoolYear.Items.Add("" + (i + j));
                }

            }

            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");
            circularProgress1.Hide();

            if (_ConfigureList.Count > 0)
            {
                cboConfigure.SelectedIndex = 0;
            }
            else
            {
                cboConfigure.SelectedIndex = -1;
            }

            if (_Configure.PrintAttendanceList == null)
                _Configure.PrintAttendanceList = new List<string>();

            DataGridViewTextBoxColumn colName = new DataGridViewTextBoxColumn();
            colName.HeaderText = "節次分類";
            colName.MinimumWidth = 70;
            colName.Name = "colName";
            colName.ReadOnly = true;
            colName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            colName.Width = 70;
            this.dgAttendanceData.Columns.Add(colName);

            List<string> colNameList = new List<string>();
            foreach (string absence in absenceList)
            {
                System.Windows.Forms.DataGridViewCheckBoxColumn newCol = new DataGridViewCheckBoxColumn();
                newCol.HeaderText = absence;
                newCol.Width = 55;
                newCol.ReadOnly = false;
                newCol.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
                newCol.Tag = absence;
                colNameList.Add(absence);
                newCol.ValueType = typeof(bool);
                this.dgAttendanceData.Columns.Add(newCol);
            }

            foreach (string str in typeList)
            {
                int rowIdx = dgAttendanceData.Rows.Add();
                dgAttendanceData.Rows[rowIdx].Tag = str;
                dgAttendanceData.Rows[rowIdx].Cells[0].Value = str;
                int colIdx = 1;

                foreach (string str1 in colNameList)
                {
                    string key = str + "_" + str1;
                    DataGridViewCheckBoxCell cell = new DataGridViewCheckBoxCell();
                    cell.Tag = key;
                    cell.Value = false;
                    if (_Configure.PrintAttendanceList.Contains(key))
                        cell.Value = true;

                    dgAttendanceData.Rows[rowIdx].Cells[colIdx] = cell;
                    colIdx++;
                }
            }

            string userSelectConfigName = "";
            // 檢查畫面上是否有使用者選的
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    userSelectConfigName = conf.Name;
                    break;
                }

            if (!string.IsNullOrEmpty(_Configure.SelSetConfigName))
                cboConfigure.Text = userSelectConfigName;


            // 不排名學生類別項目放入
            cboNotRankedFilter.Items.Add("");
            foreach (string name in _StudTagDict.Keys)
                cboNotRankedFilter.Items.Add(name);

            if (!string.IsNullOrEmpty(_Configure.NotRankedTagNameFilter))
            {
                if (cboNotRankedFilter.Items.Contains(_Configure.NotRankedTagNameFilter))
                    cboNotRankedFilter.Text = _Configure.NotRankedTagNameFilter;
            }
            btnSaveConfig.Enabled = btnPrint.Enabled = true;
        }

        void bkw_DoWork(object sender, DoWorkEventArgs e)
        {
            bkw.ReportProgress(1);           
        

            // 檢查預設樣板是否存在
            _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);

            // 沒有設定檔，建立預設設定檔
            if (_UDTConfigList.Count < 2)
            {
                bkw.ReportProgress(10);
                foreach (string name in Global.DefaultConfigNameList())
                {
                    Configure cn = new Configure();
                    cn.Name = name;
                    cn.SchoolYear = K12.Data.School.DefaultSchoolYear;
                    cn.Semester = K12.Data.School.DefaultSemester;
                    DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                    conf.Name = name;
                    conf.UDTTableName = Global._UDTTableName;
                    conf.ProjectName = Global._ProjectName;
                    conf.Type = Global._DefaultConfTypeName;
                    _UDTConfigList.Add(conf);

                    //// 設預設樣板
                    //switch (name)
                    //{
                    //    case "領域成績單":
                    //        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_領域成績單));
                    //        break;

                    //    case "科目成績單":
                    //        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目成績單));
                    //        break;

                    //    case "科目及領域成績單_領域組距":
                    //        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_領域組距));
                    //        break;
                    //    case "科目及領域成績單_科目組距":
                    //        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹_科目及領域成績單_科目組距));
                    //        break;                      
                    //}

                    //if (cn.Template == null)
                        cn.Template = new Document(new MemoryStream(Properties.Resources.新竹學期成績單樣板));
                    cn.Encode();
                    cn.Save();
                }
                if (_UDTConfigList.Count > 0)
                    DAO.UDTTransfer.InsertConfigData(_UDTConfigList);
            }
            bkw.ReportProgress(20);
            // 取的設定資料
            _ConfigureList = _AccessHelper.Select<Configure>();

            bkw.ReportProgress(40);
            // 缺曠資料
            foreach (JHPeriodMappingInfo info in JHPeriodMapping.SelectAll())
            {
                if (!typeList.Contains(info.Type))
                    typeList.Add(info.Type);
            }

            bkw.ReportProgress(70);
            foreach (JHAbsenceMappingInfo info in JHAbsenceMapping.SelectAll())
            {
                if (!absenceList.Contains(info.Name))
                    absenceList.Add(info.Name);
            }
            bkw.ReportProgress(80);
            // 所有有學生類別
            _StudTagDict = Utility.GetStudentTagRefDict();
           

            bkw.ReportProgress(100);
        }

        void _bgWorkReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績報表產生中...", e.ProgressPercentage);
        }

        void _bgWorkReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                btnSaveConfig.Enabled = true;
                btnPrint.Enabled = true;

                if (_ErrorList.Count > 0)
                {
                    StringBuilder sb = new StringBuilder();
                    //sb.AppendLine("樣板內科目合併欄位不足，請新增：");
                    //sb.AppendLine(string.Join(",", _ErrorList.ToArray()));
                    sb.AppendLine("1.樣板內科目合併欄位不足，請檢查樣板。");
                    sb.AppendLine("2.如果使用只有領域樣板，請忽略此訊息。");
                    if(_ErrorDomainNameList.Count>0)
                        sb.AppendLine(string.Join(",",_ErrorDomainNameList.ToArray()));

                    FISCA.Presentation.Controls.MsgBox.Show(sb.ToString(), "樣板內科目合併欄位不足", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                }

                FISCA.Presentation.MotherForm.SetStatusBarMessage("學期成績報表產生完成");
                System.Diagnostics.Process.Start(pathW);
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
            }
        }

        void _bgWorkReport_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 資料讀取
            _bgWorkReport.ReportProgress(1);

            // 每次合併後放入，最後再合成一張
            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.新竹學期成績單樣板));

            _ErrorList.Clear();
            _ErrorDomainNameList.Clear();
            _TemplateSubjectNameList.Clear();

            // 取得樣板內科目名稱
            foreach (string name in docTemplate.MailMerge.GetFieldNames())
            {
                if (name.Contains("科目名稱"))
                    _TemplateSubjectNameList.Add(name);
            }

            // 校名
            string SchoolName = K12.Data.School.ChineseName;
            // 校長
            string ChancellorChineseName = JHSchool.Data.JHSchoolInfo.ChancellorChineseName;
            // 教務主任
            string EduDirectorName = JHSchool.Data.JHSchoolInfo.EduDirectorName;

            // 班級
            Dictionary<string, ClassRecord> ClassDict = new Dictionary<string, ClassRecord>();
            foreach (ClassRecord cr in Class.SelectAll())
                ClassDict.Add(cr.ID, cr);

            // 不排名學生ID
            List<string> notRankStudIDList = new List<string> ();
            if (_StudTagDict.ContainsKey(_SelNotRankedFilter))
                notRankStudIDList.AddRange(_StudTagDict[_SelNotRankedFilter].ToArray());

            // 移除所選學生不排名
            foreach (string id in notRankStudIDList)
                _StudentIDList.Remove(id);

            // 所選學生資料
            List<StudentRecord> StudRecList = Student.SelectByIDs(_StudentIDList);

            // 班級年級區分,沒有年級不處理
            Dictionary<int, List<StudentRecord>> studGradeDict = new Dictionary<int, List<StudentRecord>>();
            List<string> studIDAllList =new List<string> ();
            foreach (StudentRecord studRec in Student.SelectAll())
            {
                // 不排名學生ID
                if (notRankStudIDList.Contains(studRec.ID))
                    continue;

                if (studRec.Status == StudentRecord.StudentStatus.一般)
                {
                    if (ClassDict.ContainsKey(studRec.RefClassID))
                    {
                        if (ClassDict[studRec.RefClassID].GradeYear.HasValue)
                        {
                            int gr = ClassDict[studRec.RefClassID].GradeYear.Value;

                            if (!studGradeDict.ContainsKey(gr))
                                studGradeDict.Add(gr, new List<StudentRecord>());

                            studIDAllList.Add(studRec.ID);
                            studGradeDict[gr].Add(studRec);
                        }
                    }

                }
            }
            _bgWorkReport.ReportProgress(15);
            
            
            // 取得所選學生該學年度學期學期成績
            Dictionary<string, JHSemesterScoreRecord> SelectStudSemesterRecordDict = new Dictionary<string, JHSemesterScoreRecord>();
            List<JHSemesterScoreRecord> selectStudSemsScoreList = new List<JHSemesterScoreRecord>();
            _StudentScoreDict.Clear();
            if (_StudentIDList.Count > 0 && _SelSchoolYear > 0 && _SelSemester > 0)
            {
                selectStudSemsScoreList = JHSemesterScore.SelectByStudentIDs(_StudentIDList);
                foreach (JHSemesterScoreRecord rec in selectStudSemsScoreList)
                {                    
                    if (rec.SchoolYear == _SelSchoolYear && rec.Semester == _SelSemester)
                    {
                        if (!SelectStudSemesterRecordDict.ContainsKey(rec.RefStudentID))
                        {
                            SelectStudSemesterRecordDict.Add(rec.RefStudentID, rec);

                            DAO.StudentScore ss = new DAO.StudentScore();
                            ss.StudentID = rec.RefStudentID;
                            ss.LearnDomainScore = rec.LearnDomainScore;
                            ss.CourseLearnScore = rec.CourseLearnScore;
                            
                            // 領域
                            foreach (DomainScore ds in rec.Domains.Values)
                            {
                                ss.DomainScoreDict.Add(ds.Domain, ds);
                                if(!string.IsNullOrWhiteSpace(ds.Text))
                                    ss.DomainTextList.Add(ds.Text);
                            }
                            // 科目
                            foreach (SubjectScore sss in rec.Subjects.Values)
                            {
                                // 有成績才放入
                                if (!_SelSubjNameList.Contains(sss.Subject))
                                    continue;

                                string dName = sss.Domain;
                                if (string.IsNullOrEmpty(dName))
                                    dName = "彈性課程";

                                if (!ss.SubjectScoreDict.ContainsKey(dName))
                                    ss.SubjectScoreDict.Add(dName, new List<SubjectScore>());

                                ss.SubjectScoreDict[dName].Add(sss);

                                if(!string.IsNullOrWhiteSpace(sss.Text))
                                    ss.SubjecTextList.Add(sss.Text);
                            }
                            _StudentScoreDict.Add(rec.RefStudentID, ss);
                        }
                    }
                }            
            }


            // 取得全年級學生該學年度學期學期成績
            Dictionary<string, JHSemesterScoreRecord> AllStudSemesterRecordDict = new Dictionary<string, JHSemesterScoreRecord>();
            List<JHSemesterScoreRecord> allStudSemsScoreList = new List<JHSemesterScoreRecord>();
            if (studIDAllList.Count > 0 && _SelSchoolYear > 0 && _SelSemester > 0)
            {
                allStudSemsScoreList=JHSemesterScore.SelectByStudentIDs(studIDAllList);
                foreach (JHSemesterScoreRecord rec in allStudSemsScoreList)
                {
                    if (rec.SchoolYear == _SelSchoolYear && rec.Semester == _SelSemester)
                    {
                        if (!AllStudSemesterRecordDict.ContainsKey(rec.RefStudentID))
                            AllStudSemesterRecordDict.Add(rec.RefStudentID, rec);
                    }
                }        
            }

            #region 領域成績區間處理
            // 處理領域成績區間
            // 班級
            Dictionary<string, Dictionary<string, DAO.DomainRangeCount>> dmRClasssDict = new Dictionary<string, Dictionary<string, DAO.DomainRangeCount>>();
            // 年級
            Dictionary<int, Dictionary<string, DAO.DomainRangeCount>> dmGradeDict = new Dictionary<int, Dictionary<string, DAO.DomainRangeCount>>();

            List<string> dmClassDNList = new List<string>();
            List<string> dmGradeDNList = new List<string>();

            foreach (int gr in studGradeDict.Keys)
            {
                foreach (StudentRecord rec in studGradeDict[gr])
                {                   
                    // 有成績
                    if (AllStudSemesterRecordDict.ContainsKey(rec.ID))
                    {
                        foreach (DomainScore ds in AllStudSemesterRecordDict[rec.ID].Domains.Values)
                        {                            
                            // 領域班組距名稱
                            if (!dmClassDNList.Contains(ds.Domain))
                                dmClassDNList.Add(ds.Domain);

                            // 領域年組距名稱
                            if (!dmGradeDNList.Contains(ds.Domain))
                                dmGradeDNList.Add(ds.Domain);

                            // 放入領域成績至班組距
                            if (rec.RefClassID != "")
                            {                                
                                if (!dmRClasssDict.ContainsKey(rec.RefClassID))
                                    dmRClasssDict.Add(rec.RefClassID, new Dictionary<string, DAO.DomainRangeCount>());

                                if (!dmRClasssDict[rec.RefClassID].ContainsKey(ds.Domain))
                                {
                                    dmRClasssDict[rec.RefClassID].Add(ds.Domain, new DAO.DomainRangeCount());
                                    dmRClasssDict[rec.RefClassID][ds.Domain].Name = ds.Domain;
                                }
                                dmRClasssDict[rec.RefClassID][ds.Domain].AddScore(ds.Score);                                
                            }

                            // 放入領域成績至年組距
                            if(!dmGradeDict.ContainsKey(gr))
                                dmGradeDict.Add(gr,new Dictionary<string,DAO.DomainRangeCount> ());

                            if (!dmGradeDict[gr].ContainsKey(ds.Domain))
                            {
                                dmGradeDict[gr].Add(ds.Domain, new DAO.DomainRangeCount());
                                dmGradeDict[gr][ds.Domain].Name = ds.Domain;
                            }
                            dmGradeDict[gr][ds.Domain].AddScore(ds.Score);
                        }
                    }
                
                }
            }
            
            #endregion
            
            _bgWorkReport.ReportProgress(45);
            
            #region 科目成績區間處理
            // 處理領域成績區間
            // 班級
            Dictionary<string, Dictionary<string, DAO.SubjectRangeCount>> subjRClasssDict = new Dictionary<string, Dictionary<string, DAO.SubjectRangeCount>>();
            // 年級
            Dictionary<int, Dictionary<string, DAO.SubjectRangeCount>> subjGradeDict = new Dictionary<int, Dictionary<string, DAO.SubjectRangeCount>>();

            List<string> subjClassDNList = new List<string>();
            List<string> subjGradeDNList = new List<string>();

            foreach (int gr in studGradeDict.Keys)
            {
                foreach (StudentRecord rec in studGradeDict[gr])
                {
                    // 有成績
                    if (AllStudSemesterRecordDict.ContainsKey(rec.ID))
                    {
                        foreach (SubjectScore ss in AllStudSemesterRecordDict[rec.ID].Subjects.Values)
                        {
                            // 有勾選才放入
                            if (!_SelSubjNameList.Contains(ss.Subject))
                                continue;

                            // 領域班組距名稱
                            if (!subjClassDNList.Contains(ss.Subject))
                                subjClassDNList.Add(ss.Subject);

                            // 領域年組距名稱
                            if (!subjGradeDNList.Contains(ss.Subject))
                                subjGradeDNList.Add(ss.Subject);

                            // 放入領域成績至班組距
                            if (rec.RefClassID != "")
                            {
                                if (!subjRClasssDict.ContainsKey(rec.RefClassID))
                                    subjRClasssDict.Add(rec.RefClassID, new Dictionary<string, DAO.SubjectRangeCount>());

                                if (!subjRClasssDict[rec.RefClassID].ContainsKey(ss.Subject))
                                {
                                    subjRClasssDict[rec.RefClassID].Add(ss.Subject, new DAO.SubjectRangeCount());
                                    subjRClasssDict[rec.RefClassID][ss.Subject].Name = ss.Subject;
                                }
                                subjRClasssDict[rec.RefClassID][ss.Subject].AddScore(ss.Score);
                            }

                            // 放入領域成績至年組距
                            if (!subjGradeDict.ContainsKey(gr))
                                subjGradeDict.Add(gr, new Dictionary<string, DAO.SubjectRangeCount>());

                            if (!subjGradeDict[gr].ContainsKey(ss.Subject))
                            {
                                subjGradeDict[gr].Add(ss.Subject, new DAO.SubjectRangeCount());
                                subjGradeDict[gr][ss.Subject].Name = ss.Subject;
                            }
                            subjGradeDict[gr][ss.Subject].AddScore(ss.Score);
                        }
                    }

                }
            }
            
            #endregion             
            
            #region 日常生活表現處理

            // 取得日常生活表現處理設定
            Dictionary<string,string> DLBehaviorDict=Utility.GetDLBehaviorConfigNameDict();

            // 取得學生日常生活表現
            _StudTextScoreXMLDict = Utility.GetStudentTextScoreDict(_StudentIDList, _SelSchoolYear, _SelSemester);
            
            #endregion

            // 缺曠資料區間統計
             _AttendanceDict = Utility.GetAttendanceCountByDate(StudRecList, _BeginDate, _EndDate);

            // 獎懲資料
            Dictionary<string, Dictionary<string, int>> DisciplineCountDict = Utility.GetDisciplineCountByDate(_StudentIDList, _BeginDate, _EndDate);

            // 服務學習
            Dictionary<string, decimal> ServiceLearningDict = Utility.GetServiceLearningDetailByDate(_StudentIDList, _BeginDate, _EndDate);

            // 領域組距
            List<string> li2 = new List<string>();
            li2.Add("R100_u");
            li2.Add("R90_99");
            li2.Add("R80_89");
            li2.Add("R70_79");
            li2.Add("R60_69");
            li2.Add("R50_59");
            li2.Add("R40_49");
            li2.Add("R30_39");
            li2.Add("R20_29");
            li2.Add("R10_19");
            li2.Add("R0_9");

            #endregion


            _bgWorkReport.ReportProgress(60);

            List<string> domainLi = new List<string>();

            List<string> subjLi = new List<string>();
            subjLi.Add("科目名稱");
            subjLi.Add("科目權數");
            subjLi.Add("科目成績");
            subjLi.Add("科目等第");
            

            List<string> subjColList = new List<string>();
            foreach (string dName in Global.DomainNameList())
            {
                for (int i = 1; i <= 7; i++)
                {
                    foreach (string sName in subjLi)
                    {
                        string key = dName + "_" + sName +i;
                        subjColList.Add(key);
                    }
                }
            }
            

            // 學生筆學期歷程
            Dictionary<string, SemesterHistoryItem> StudShiDict = new Dictionary<string, SemesterHistoryItem>();
            // 取得學期歷程，給班級、座號、班導師使用，條件：所選學生
            List<SemesterHistoryRecord> SemesterHistoryRecordList = SemesterHistory.SelectByStudentIDs(_StudentIDList);
            foreach (SemesterHistoryRecord shr in SemesterHistoryRecordList)
            {
                foreach(SemesterHistoryItem shi in shr.SemesterHistoryItems)
                if (shi.SchoolYear == _SelSchoolYear && shi.Semester == _SelSemester)
                {
                    if (!StudShiDict.ContainsKey(shi.RefStudentID))
                        StudShiDict.Add(shi.RefStudentID, shi);
                }
            }
            
            #region 處理合併 DataTable 相關資料
            // 儲存資料用 Data Table
            DataTable dt = new DataTable();  
            Document doc = new Document();
            DataTable dtAtt = new DataTable();
            List<Document> docList = new List<Document>();

            string BehaviorName = "";
            string OtherRecommendName = "";
            string DailyLifeRecommendName = "";

            if (DLBehaviorDict.ContainsKey("日常行為表現"))
                BehaviorName = DLBehaviorDict["日常行為表現"];

            if (DLBehaviorDict.ContainsKey("其它表現"))
                OtherRecommendName = DLBehaviorDict["其它表現"];

            if (DLBehaviorDict.ContainsKey("綜合評語"))
                DailyLifeRecommendName = DLBehaviorDict["綜合評語"];

            List<string> DLStringList = new List<string>();
            DLStringList.Add("愛整潔");
            DLStringList.Add("有禮貌");
            DLStringList.Add("守秩序");
            DLStringList.Add("責任心");
            DLStringList.Add("公德心");
            DLStringList.Add("友愛關懷");
            DLStringList.Add("團隊合作");
            DLStringList.Add("其他表現");
            DLStringList.Add("綜合評語");


            // 填值
            foreach (StudentRecord StudRec in StudRecList)
            {
                #region 處理DataTable欄位名稱

                DataRow row = dt.NewRow();
                dtAtt.Columns.Clear();
                dtAtt.Clear();
                dt.Clear();
                dt.Columns.Clear();

                dtAtt.Columns.Add("缺曠紀錄");
                DataRow rowT = dtAtt.NewRow();

                // 取得欄位
                foreach (string colName in Global.DTColumnsList())
                    dt.Columns.Add(colName);

                // 組距欄位
                // 領域班級組距
                foreach (string n1 in dmClassDNList)
                {
                    foreach (string n2 in li2)
                    {
                        string colName = "班級_" + n1 + "_" + n2;
                        dt.Columns.Add(colName);
                    }
                }
                // 年級
                foreach (string n1 in dmGradeDNList)
                {
                    foreach (string n2 in li2)
                    {
                        string colName = "年級_" + n1 + "_" + n2;
                        dt.Columns.Add(colName);
                    }
                }

                // 科目
                // 領域班級組距
                Dictionary<string, string> colSubjMapDict = new Dictionary<string, string>();
                int c1 = 1, g1 = 1;
                foreach (string n1 in subjClassDNList)
                {
                    string sName1 = "s班級_科目名稱" + n1;
                    string sName2 = "s班級_科目名稱" + c1;
                    dt.Columns.Add(sName2);
                    // 放入科目名稱
                    colSubjMapDict.Add(sName1, sName2);
                    foreach (string n2 in li2)
                    {
                        string colName = "s班級_" + n1 + "_" + n2;
                        string colVal = "s班級_" + "科目" + c1 + "_" + n2;
                        colSubjMapDict.Add(colName, colVal);
                        dt.Columns.Add(colVal);
                    }
                    c1++;
                }
                // 年級
                foreach (string n1 in subjGradeDNList)
                {
                    string gName1 = "s年級_科目名稱" + n1;
                    string gName2 = "s年級_科目名稱" + g1;
                    dt.Columns.Add(gName2);
                    colSubjMapDict.Add(gName1, gName2);
                    foreach (string n2 in li2)
                    {
                        string colName = "s年級_" + n1 + "_" + n2;
                        string colVal = "s年級_" + "科目" + g1 + "_" + n2;
                        colSubjMapDict.Add(colName, colVal);
                        dt.Columns.Add(colVal);
                    }
                    g1++;
                }

                // 新增科目成績欄位
                foreach (string colName in subjColList)
                    dt.Columns.Add(colName);

                // 新增領域成績欄位
                foreach (string dName in Global.DomainNameList())
                {
                    dt.Columns.Add(dName + "_領域權數");
                    dt.Columns.Add(dName + "_領域成績");
                    dt.Columns.Add(dName + "_領域等第");
                }
                
                // 處理日常生活表現欄位
                // 日常行為表現:DailyBehavior,Item
                // 日常行為表現_愛整潔
                string strDailyBehavior = "日常行為表現";
                if (Utility.DLBehaviorConfigItemNameDict.ContainsKey(strDailyBehavior))
                {
                    foreach (string str2 in Utility.DLBehaviorConfigItemNameDict[strDailyBehavior])
                        dt.Columns.Add(BehaviorName+ "_" + str2);
                }
                // 其他表現:OtherRecommend
                dt.Columns.Add(OtherRecommendName);
                // 綜合評語:DailyLifeRecommend
                dt.Columns.Add(DailyLifeRecommendName);

                foreach (string str in DLStringList)
                {
                    if (!dt.Columns.Contains(str))
                        dt.Columns.Add(str);
                }

                #endregion
                
                dt.TableName = StudRec.ID;
                row["StudentID"] = StudRec.ID;
                row["學校名稱"] = SchoolName;
                row["學年度"] = _SelSchoolYear;
                row["學期"] = _SelSemester;
                

                // 班級、座號、班導師 使用學期歷程內
                if (StudShiDict.ContainsKey(StudRec.ID))
                {
                    row["班級"] = StudShiDict[StudRec.ID].ClassName;
                    row["座號"] = StudShiDict[StudRec.ID].SeatNo;
                    row["班導師"] = StudShiDict[StudRec.ID].Teacher;
                }

                row["學號"] = StudRec.StudentNumber;
                row["姓名"] = StudRec.Name;


                // 傳入 ID當 Key
                // row["缺曠紀錄"] = StudRec.ID;
                rowT["缺曠紀錄"] = StudRec.ID;
                // 獎懲區間統計值
                if (DisciplineCountDict.ContainsKey(StudRec.ID))
                {
                    foreach (string str in Global.GetDisciplineNameList())
                    {
                        string key = str + "區間統計";
                        if (DisciplineCountDict[StudRec.ID].ContainsKey(str))
                            row[key] = DisciplineCountDict[StudRec.ID][str];
                    }
                }              

                // 處理領域組距相關
                // 班級
                string kClassKey = "";

                List<DAO.DomainRangeCount.DomainRangeType> dtypeList = new List<DAO.DomainRangeCount.DomainRangeType>();
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R100_u);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R90_99);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R80_89);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R70_79);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R60_69);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R50_59);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R40_49);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R30_39);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R20_29);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R10_19);
                dtypeList.Add(DAO.DomainRangeCount.DomainRangeType.R0_9);

                if (dmRClasssDict.ContainsKey(StudRec.RefClassID))
                {
                    foreach (KeyValuePair<string, DAO.DomainRangeCount> data in dmRClasssDict[StudRec.RefClassID])
                    {
                        foreach (DAO.DomainRangeCount.DomainRangeType dtType in dtypeList)
                        {
                            kClassKey = "班級_" + data.Key + "_" + dtType.ToString();
                            row[kClassKey] = data.Value.GetRankCount(dtType);
                        }
                    }
                }

                // 年級
                int grY = 0;
                if (ClassDict.ContainsKey(StudRec.RefClassID))
                    if (ClassDict[StudRec.RefClassID].GradeYear.HasValue)
                        grY = ClassDict[StudRec.RefClassID].GradeYear.Value;

                string kGradeKey = "";
                if (dmGradeDict.ContainsKey(grY))
                {
                    foreach (KeyValuePair<string, DAO.DomainRangeCount> data in dmGradeDict[grY])
                    {
                        foreach (DAO.DomainRangeCount.DomainRangeType dtType in dtypeList)
                        {
                            kGradeKey = "年級_" + data.Key + "_" + dtType.ToString();
                            row[kGradeKey] = data.Value.GetRankCount(dtType);
                        }
                    }
                }

                // 處理科目組距相關
                // 班級
                string sClassKey = "";

                List<DAO.SubjectRangeCount.SubjectRangeType> stypeList = new List<DAO.SubjectRangeCount.SubjectRangeType>();
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R100_u);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R90_99);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R80_89);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R70_79);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R60_69);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R50_59);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R40_49);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R30_39);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R20_29);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R10_19);
                stypeList.Add(DAO.SubjectRangeCount.SubjectRangeType.R0_9);

                if (subjRClasssDict.ContainsKey(StudRec.RefClassID))
                {
                    foreach (KeyValuePair<string, DAO.SubjectRangeCount> data in subjRClasssDict[StudRec.RefClassID])
                    {
                        string ssKey = "s班級_科目名稱" + data.Key;
                        if (colSubjMapDict.ContainsKey(ssKey))
                            row[colSubjMapDict[ssKey]] = data.Key;

                        foreach (DAO.SubjectRangeCount.SubjectRangeType dtType in stypeList)
                        {
                            sClassKey = "s班級_" + data.Key + "_" + dtType.ToString();
                            if (colSubjMapDict.ContainsKey(sClassKey))
                            {                 
                                row[colSubjMapDict[sClassKey]] = data.Value.GetRankCount(dtType);
                            }
                        }
                    }
                }

                // 年級
                int sgrY = 0;
                if (ClassDict.ContainsKey(StudRec.RefClassID))
                    if (ClassDict[StudRec.RefClassID].GradeYear.HasValue)
                        sgrY = ClassDict[StudRec.RefClassID].GradeYear.Value;

                string sGradeKey = "";
                if (subjGradeDict.ContainsKey(sgrY))
                {
                    foreach (KeyValuePair<string, DAO.SubjectRangeCount> data in subjGradeDict[sgrY])
                    {
                        string ssKey = "s年級_科目名稱" + data.Key;
                        if (colSubjMapDict.ContainsKey(ssKey))
                            row[colSubjMapDict[ssKey]] = data.Key;


                        foreach (DAO.SubjectRangeCount.SubjectRangeType dtType in stypeList)
                        {
                            sGradeKey = "s年級_" + data.Key + "_" + dtType.ToString();
                            if (colSubjMapDict.ContainsKey(sGradeKey))
                            {                             
                                row[colSubjMapDict[sGradeKey]] = data.Value.GetRankCount(dtType);
                            }
                        }
                    }
                }

                row["服務學習時數"] = "";
                if (ServiceLearningDict.ContainsKey(StudRec.ID))
                    row["服務學習時數"] = ServiceLearningDict[StudRec.ID];

                row["校長"] = ChancellorChineseName;
                row["教務主任"] = EduDirectorName;
                row["區間開始日期"] = _BeginDate.ToShortDateString();
                row["區間結束日期"] = _EndDate.ToShortDateString();
                row["成績校正日期"] = _ScoreEditDate;

                // 處理成績
                if (_StudentScoreDict.ContainsKey(StudRec.ID))
                {
                    DAO.StudentScore ss = _StudentScoreDict[StudRec.ID];
                    
                    // 領域
                    foreach (string dName in Global.DomainNameList())
                    {
                        if (ss.DomainScoreDict.ContainsKey(dName))
                        {
                            if(ss.DomainScoreDict[dName].Credit.HasValue)
                                row[dName + "_領域權數"] = ss.DomainScoreDict[dName].Credit.Value;

                            if (ss.DomainScoreDict[dName].Score.HasValue)
                            {
                                row[dName + "_領域成績"] = ss.DomainScoreDict[dName].Score.Value;
                                row[dName + "_領域等第"] = _degreeMapper.GetDegreeByScore(ss.DomainScoreDict[dName].Score.Value);
                            }                            
                        }
                    }

                    // 科目(依領域分科目)
                    foreach (string dName in Global.DomainNameList())
                    {
                        int subjCot = 1;
                        if (ss.SubjectScoreDict.ContainsKey(dName))
                        {
                            foreach (SubjectScore s in ss.SubjectScoreDict[dName])
                            {
                                row[dName+"_科目名稱" + subjCot] = s.Subject;
                                if (s.Credit.HasValue)
                                    row[dName+"_科目權數" + subjCot] = s.Credit.Value;
                                if (s.Score.HasValue)
                                {
                                    row[dName+"_科目成績" + subjCot] = s.Score.Value;
                                    row[dName+"_科目等第" + subjCot] = _degreeMapper.GetDegreeByScore(s.Score.Value);
                                }
                                subjCot++;
                            }
                        }                    
                    }

                    // 領域文字描述
                    row["領域文字描述"] = string.Join(",", ss.DomainTextList.ToArray());

                    // 科目文字描述
                    row["科目文字描述"] = string.Join(",", ss.SubjecTextList.ToArray());

                    // 學習領域成績
                    if (ss.LearnDomainScore.HasValue)
                    {
                        row["學習領域成績"] = ss.LearnDomainScore.Value;
                        row["學習領域等第"] = _degreeMapper.GetDegreeByScore(ss.LearnDomainScore.Value);
                    }
                    // 課程學習領域成績
                    if (ss.CourseLearnScore.HasValue)
                    {
                        row["課程學習領域成績"] = ss.CourseLearnScore.Value;
                        row["課程學習領域等第"] = _degreeMapper.GetDegreeByScore(ss.CourseLearnScore.Value);
                    }
                }

                // 處理日常生活表現
                if (_StudTextScoreXMLDict.ContainsKey(StudRec.ID))
                {
                    DAO.StudTextScoreXML studXml = _StudTextScoreXMLDict[StudRec.ID];
                    // DailyBehavior:日常行為表現_愛整潔
                    if (Utility.DLBehaviorConfigItemNameDict.ContainsKey(strDailyBehavior))
                        foreach (string str2 in Utility.DLBehaviorConfigItemNameDict[strDailyBehavior])
                        {
                            row[BehaviorName + "_" + str2] = studXml.GetDailyBehavior(strDailyBehavior, str2);
                            if (DLStringList.Contains(str2))
                                row[str2] = studXml.GetDailyBehavior(strDailyBehavior, str2);
                        }
                    //OtherRecommend:其他表現
                    row[OtherRecommendName] = studXml.GetOtherRecommend(OtherRecommendName);
                    row["其他表現"] = studXml.GetOtherRecommend(OtherRecommendName);

                    //DailyLifeRecommend:綜合評語
                    row[DailyLifeRecommendName] = studXml.GetDailyLifeRecommend(DailyLifeRecommendName);
                    row["綜合評語"] = studXml.GetDailyLifeRecommend(DailyLifeRecommendName);
                }
                dt.Rows.Add(row);
                dtAtt.Rows.Add(rowT);

                // 處理固定欄位對應
                Document doc1 = new Document();
                doc1.Sections.Clear();

                // 處理動態處理(缺曠)
                Document docAtt = new Document();
                docAtt.Sections.Clear();
                docAtt.Sections.Add(docAtt.ImportNode(docTemplate.Sections[0], true));

                _builder = new DocumentBuilder(docAtt);
                docAtt.MailMerge.MergeField += new Aspose.Words.Reporting.MergeFieldEventHandler(MailMerge_MergeField);
                docAtt.MailMerge.Execute(dtAtt);

                doc1.Sections.Add(doc1.ImportNode(docAtt.Sections[0], true));
                doc1.MailMerge.Execute(dt);
                doc1.MailMerge.RemoveEmptyParagraphs = true;
                doc1.MailMerge.DeleteFields();
                docList.Add(doc1);
            }

            _bgWorkReport.ReportProgress(80);
            // debug 用           
            string ssStr = Application.StartupPath + "\\dt_debug.xml";
            dt.WriteXml(ssStr);

            #endregion

            #region Word 合併列印

            doc.Sections.Clear();
            foreach(Document doc1 in docList)
                doc.Sections.Add(doc.ImportNode(doc1.Sections[0], true));

            string reportNameW = "新竹學期成績單";
                pathW = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", "");
                if (!Directory.Exists(pathW))
                Directory.CreateDirectory(pathW);
                pathW = Path.Combine(pathW, reportNameW + ".doc");

                if (File.Exists(pathW))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPathW = Path.GetDirectoryName(pathW) + "\\" + Path.GetFileNameWithoutExtension(pathW) + (i++) + Path.GetExtension(pathW);
                        if (!File.Exists(newPathW))
                        {
                            pathW = newPathW;
                            break;
                        }
                    }
                }

                try
                {
                    doc.Save(pathW, Aspose.Words.SaveFormat.Doc);

                }
                catch (Exception exow)
                {

                }
            doc = null;
            docList.Clear();

            GC.Collect();
            #endregion
            _bgWorkReport.ReportProgress(100);
        }

        void MailMerge_MergeField(object sender, Aspose.Words.Reporting.MergeFieldEventArgs e)
        {
            if (e.FieldName == "缺曠紀錄")
            {
                if (_builder.MoveToMergeField(e.FieldName))
                {
                    string sid = e.FieldValue.ToString();

                    Dictionary<string, int> dataDict = new Dictionary<string, int>();    
                        List<string> colNameList = new List<string>();
                        if (_AttendanceDict.ContainsKey(sid))
                            dataDict = _AttendanceDict[sid];
                        //dataDict.Keys
                        
                        foreach (string name in _SelAttendanceList)
                            colNameList.Add(name.Replace("_",""));

                        //colNameList.Sort();
                        int colCount=colNameList.Count;

                        if (colCount > 0)
                        {
                            Cell cell = _builder.CurrentParagraph.ParentNode as Cell;
                            cell.CellFormat.LeftPadding = 0;
                            cell.CellFormat.RightPadding = 0;
                            double width = cell.CellFormat.Width;
                            int columnCount = colCount;
                            double miniUnitWitdh = width / (double)columnCount;

                            Table table = _builder.StartTable();

                            //(table.ParentNode.ParentNode as Row).RowFormat.LeftIndent = 0;
                            double p = _builder.RowFormat.LeftIndent;
                            _builder.RowFormat.HeightRule = HeightRule.Exactly;
                            _builder.RowFormat.Height = 18.0;
                            _builder.RowFormat.LeftIndent = 0;

                            // 缺曠名稱
                            foreach (string name in colNameList)
                            {
                                Cell c1 = _builder.InsertCell();
                                c1.CellFormat.Width = miniUnitWitdh;
                                c1.CellFormat.WrapText = true;
                                _builder.Write(name);                            
                            }
                            _builder.EndRow();

                            // 缺曠統計
                            foreach (string name in colNameList)
                            {
                                Cell c1 = _builder.InsertCell();
                                c1.CellFormat.Width = miniUnitWitdh;
                                c1.CellFormat.WrapText = true;
                                if (dataDict.ContainsKey(name)) 
                                    _builder.Write(dataDict[name].ToString()); 
                                else
                                    _builder.Write("");
                            }
                            _builder.EndRow();                            

                            _builder.EndTable();

                            //去除表格四邊的線
                            foreach (Cell c in table.FirstRow.Cells)
                                c.CellFormat.Borders.Top.LineStyle = LineStyle.None;

                            foreach (Cell c in table.LastRow.Cells)
                                c.CellFormat.Borders.Bottom.LineStyle = LineStyle.None;

                            foreach (Row r in table.Rows)
                            {
                                r.FirstCell.CellFormat.Borders.Left.LineStyle = LineStyle.None;
                                r.LastCell.CellFormat.Borders.Right.LineStyle = LineStyle.None;
                            }

                            _builder.RowFormat.LeftIndent = p;
                        }
                  
                }
            }            
        }


        // 載入學生所屬學年度學習的試別，科目，並排序
        private void LoadSemesterSubject()
        {
            // 取得該學年度學期所有學生的試別修課科目
            _SelSchoolYear = _SelSemester = 0;
            int ss, sc;
            if (int.TryParse(cboSchoolYear.Text, out ss))
                _SelSchoolYear = ss;

            if (int.TryParse(cboSemester.Text, out sc))
                _SelSemester = sc;

            _SemesterSubjectNameList.Clear();

            List<JHSemesterScoreRecord> SemsScore = new List<JHSemesterScoreRecord>();
            
            if(_StudentIDList.Count>0 && _SelSchoolYear>0 && _SelSemester>0)
                SemsScore=JHSemesterScore.SelectBySchoolYearAndSemester(_StudentIDList, _SelSchoolYear, _SelSemester);

            // 填入學升學期科目名稱
            foreach (JHSemesterScoreRecord rec in SemsScore)
                foreach (string subjName in rec.Subjects.Keys)
                    if (!_SemesterSubjectNameList.Contains(subjName))
                        _SemesterSubjectNameList.Add(subjName);

            // 排序
            _SemesterSubjectNameList.Sort(new StringComparer("國文"
                                , "英文"
                                , "數學"
                                , "理化"
                                , "生物"
                                , "社會"
                                , "物理"
                                , "化學"
                                , "歷史"
                                , "地理"
                                , "公民"));


        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            DisSelect();
            _SelSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
            _SelSemester = int.Parse(K12.Data.School.DefaultSemester);

            bkw.RunWorkerAsync();
        }

        private void LoadSubject()
        {
            lvSubject.Items.Clear();
            foreach (string subjName in _SemesterSubjectNameList)
                lvSubject.Items.Add(subjName);
        }


        private void lnkCopyConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = _Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;                
                conf.PrintSubjectList.AddRange(_Configure.PrintSubjectList);
                conf.SchoolYear = _Configure.SchoolYear;
                conf.Semester = _Configure.Semester;
                conf.SubjectLimit = _Configure.SubjectLimit;
                conf.Template = _Configure.Template;
                conf.BeginDate = _Configure.BeginDate;
                conf.EndDate = _Configure.EndDate;
                conf.ScoreEditDate = _Configure.ScoreEditDate;
                if (conf.PrintAttendanceList == null)
                    conf.PrintAttendanceList = new List<string>();
                conf.PrintAttendanceList.AddRange(_Configure.PrintAttendanceList);
                conf.Encode();
                conf.Save();
                _ConfigureList.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }

        public Configure _Configure { get; private set; }

        private void lnkDelConfig_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;

            // 檢查是否是預設設定檔名稱，如果是無法刪除
            if (Global.DefaultConfigNameList().Contains(_Configure.Name))
            {
                FISCA.Presentation.Controls.MsgBox.Show("系統預設設定檔案無法刪除");
                return;
            }

            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _ConfigureList.Remove(_Configure);
                if (_Configure.UID != "")
                {
                    _Configure.Deleted = true;
                    _Configure.Save();
                }
                var conf = _Configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // 停用可選功能
        private void DisSelect()
        {
            cboConfigure.Enabled = false;            
            cboSchoolYear.Enabled = false;
            cboSemester.Enabled = false;
            btnSaveConfig.Enabled = false;
            btnPrint.Enabled = false;
        }

        // 啟用可選功能
        private void EnbSelect()
        {
            cboConfigure.Enabled = true;            
            cboSchoolYear.Enabled = true;
            cboSemester.Enabled = true;
            btnSaveConfig.Enabled = true;
            btnPrint.Enabled = true;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (dtBegin.IsEmpty || dtEnd.IsEmpty)
            {
                FISCA.Presentation.Controls.MsgBox.Show("日期區間必須輸入!");
                return;
            }

            if (dtBegin.Value > dtEnd.Value)
            {
                FISCA.Presentation.Controls.MsgBox.Show("開始日期必須小於或等於結束日期!!");
                return;
            }

            int sc, ss;
            if (int.TryParse(cboSchoolYear.Text, out sc))
            {
                _SelSchoolYear = sc;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學年度必填!");
                return;
            }

            if (int.TryParse(cboSemester.Text, out ss))
            {
                _SelSemester = ss;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("學期必填!");
                return;
            }

            _SelNotRankedFilter = cboNotRankedFilter.Text;
            _SelSubjNameList.Clear();

            SaveTemplate(null, null);
            
            // 使用者勾選科目
            foreach(string name in _Configure.PrintSubjectList)
                _SelSubjNameList.Add(name);
            
            // 缺曠
            foreach (string name in _Configure.PrintAttendanceList)
                _SelAttendanceList.Add(name);

            
            _BeginDate = dtBegin.Value;
            _EndDate = dtEnd.Value;

            if (dtScoreEdit.IsEmpty)
                _ScoreEditDate = "";
            else
                _ScoreEditDate = dtScoreEdit.Value.ToShortDateString();

            btnSaveConfig.Enabled = false;
            btnSaveConfig.Enabled = false;

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
            // 執行報表
            _bgWorkReport.RunWorkerAsync();
        }

        // 儲存樣板
        private void SaveTemplate(object sender, EventArgs e)
        {
            if (_Configure == null) return;
            _Configure.SchoolYear = cboSchoolYear.Text;
            _Configure.Semester = cboSemester.Text;
            _Configure.SelSetConfigName = cboConfigure.Text;
            _Configure.NotRankedTagNameFilter = _SelNotRankedFilter;
            
            // 科目
            foreach (ListViewItem item in lvSubject.Items)
            {
                if (item.Checked)
                {
                    if (!_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Add(item.Text);
                }
                else
                {
                    if (_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Remove(item.Text);
                }
            }

            if (_Configure.PrintAttendanceList == null)
                _Configure.PrintAttendanceList = new List<string>();
            
            _Configure.PrintAttendanceList.Clear();
            // 儲存缺曠選項
            foreach (DataGridViewRow drv in dgAttendanceData.Rows)
            {
                foreach (DataGridViewCell cell in drv.Cells)
                {
                    bool bl;
                    if (bool.TryParse(cell.Value.ToString(), out bl))
                    {
                        if(bl)
                            _Configure.PrintAttendanceList.Add(cell.Tag.ToString());
                    }
                }                        
            }


            foreach (ListViewItem item in lvSubject.Items)
            {
                if (item.Checked)
                {
                    if (!_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Add(item.Text);
                }
                else
                {
                    if (_Configure.PrintSubjectList.Contains(item.Text))
                        _Configure.PrintSubjectList.Remove(item.Text);
                }
            }

            // 儲存開始與結束日期
            _Configure.BeginDate = dtBegin.Value.ToShortDateString();
            _Configure.EndDate = dtEnd.Value.ToShortDateString();
            if (dtScoreEdit.IsEmpty)
                _Configure.ScoreEditDate = "";
            else
                _Configure.ScoreEditDate = dtScoreEdit.Value.ToShortDateString();

            _Configure.Encode();
            _Configure.Save();

            #region 樣板設定檔記錄用

            // 記錄使用這選的專案            
            List<DAO.UDT_ScoreConfig> uList = new List<DAO.UDT_ScoreConfig>();
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    conf.Name = cboConfigure.Text;
                    uList.Add(conf);
                    break;
                }

            if (uList.Count > 0)
            {
                DAO.UDTTransfer.UpdateConfigData(uList);
            }
            else
            {
                // 新增
                List<DAO.UDT_ScoreConfig> iList = new List<DAO.UDT_ScoreConfig>();
                DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                conf.Name = cboConfigure.Text;
                conf.ProjectName = Global._ProjectName;
                conf.Type = Global._UserConfTypeName;
                conf.UDTTableName = Global._UDTTableName;
                iList.Add(conf);
                DAO.UDTTransfer.InsertConfigData(iList);
            }
            #endregion
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisSelect();
            LoadSemesterSubject();
            LoadSubject();
            EnbSelect();
        }

        private void cboSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            DisSelect();
            LoadSemesterSubject();
            LoadSubject();
            EnbSelect();
        }

        private void cboExam_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSubject();
        }

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = btnPrint.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    
                    _Configure = new Configure();
                    _Configure.Name = dialog.ConfigName;
                    _Configure.Template = dialog.Template;
                    _Configure.SubjectLimit = dialog.SubjectLimit;
                    _Configure.SchoolYear = _DefalutSchoolYear;
                    _Configure.Semester = _DefaultSemester;
                    if(_Configure.PrintAttendanceList == null)
                        _Configure.PrintAttendanceList = new List<string>();
                    if (_Configure.PrintSubjectList == null)
                        _Configure.PrintSubjectList = new List<string>();
                   
                    _ConfigureList.Add(_Configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, _Configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;
                    _Configure.Encode();
                    _Configure.Save();
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnSaveConfig.Enabled = btnPrint.Enabled = true;
                    _Configure = _ConfigureList[cboConfigure.SelectedIndex];
                    if (_Configure.Template == null)
                        _Configure.Decode();
                    if (!cboSchoolYear.Items.Contains(_Configure.SchoolYear))
                        cboSchoolYear.Items.Add(_Configure.SchoolYear);
                    cboSchoolYear.Text = _Configure.SchoolYear;
                    cboSemester.Text = _Configure.Semester;               
                   
                    // 解析科目
                    foreach (ListViewItem lvi in lvSubject.Items)
                    {
                        if (_Configure.PrintSubjectList.Contains(lvi.Text))
                        {
                            lvi.Checked = true;
                        }
                    }
                    
                    // 解析缺曠
                    foreach (DataGridViewRow drv in dgAttendanceData.Rows)
                    {
                        foreach (DataGridViewCell cell in drv.Cells)
                        {
                            if (cell.Tag == null)
                                continue;

                            string key = cell.Tag.ToString();
                            cell.Value = false;
                            if (_Configure.PrintAttendanceList.Contains(key))
                            {
                                cell.Value = true;
                            }                        
                        }
                    }


                    // 開始與結束日期
                    DateTime dtb, dte,dtee;
                    if (DateTime.TryParse(_Configure.BeginDate, out dtb))
                        dtBegin.Value = dtb;
                    else
                        dtBegin.Value = DateTime.Now;

                    if (DateTime.TryParse(_Configure.EndDate, out dte))
                        dtEnd.Value = dte;
                    else
                        dtEnd.Value = DateTime.Now;

                    // 成績校正日期
                    if (DateTime.TryParse(_Configure.ScoreEditDate, out dtee))
                        dtScoreEdit.Value = dtee;
                    else
                        dtScoreEdit.IsEmpty = true;

                }
                else
                {
                    _Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;                                    
                  
                    // 開始與結束日期沒有預設值時給當天
                    dtBegin.Value = dtEnd.Value = DateTime.Now;                   
                }
            }
        }

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 當沒有設定檔
            if (_Configure == null) return;
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案

            string reportName = "新竹學期成績單樣板(" + _Configure.Name + ").doc";

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                _Configure.Template.Save(stream, Aspose.Words.SaveFormat.Doc);

                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.新竹學期成績單樣板, 0, Properties.Resources.新竹學期成績單樣板.Length);
                        stream.Flush();
                        stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkViewTemplate.Enabled = true;
            #endregion
        }

        private void lnkChangeTemplate_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;
            lnkChangeTemplate.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    _Configure.Template = new Aspose.Words.Document(dialog.FileName);
                    List<string> fields = new List<string>(_Configure.Template.MailMerge.GetFieldNames());
                    _Configure.SubjectLimit = 0;
                    while (fields.Contains("科目名稱" + (_Configure.SubjectLimit + 1)))
                    {
                        _Configure.SubjectLimit++;
                    }

                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
            lnkChangeTemplate.Enabled = true;
        }

        private void lnkViewMapColumns_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            lnkViewMapColumns.Enabled = false;
            Global.ExportMappingFieldWord();
            lnkViewMapColumns.Enabled = true;
        }

        private void chkSubjSelAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in lvSubject.Items)
            {
                lvi.Checked = chkSubjSelAll.Checked;
            }
        }

        private void chkAttendSelAll_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow drv in dgAttendanceData.Rows)
            {
                foreach (DataGridViewCell cell in drv.Cells)
                {
                    if(cell.ColumnIndex !=0)
                        cell.Value = chkAttendSelAll.Checked;                
                }
            }
        }

     
        
    }
}
