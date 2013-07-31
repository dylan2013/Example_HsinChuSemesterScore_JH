using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;
using K12.Data;
using System.Xml.Linq;

namespace HsinChuSemesterScore_JH
{
    public class Utility
    {

        /// <summary>
        /// 透過學生編號、開始與結束日期，取得學習服務統計值
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string,decimal> GetServiceLearningDetailByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, decimal> retVal = new Dictionary<string, decimal>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,occur_date,reason,hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') and occur_date >='" + beginDate.ToShortDateString() + "' and occur_date <='" + endDate.ToShortDateString() + "'order by ref_student_id,occur_date;";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    decimal hr;

                    string sid = dr[0].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, 0);

                    if (decimal.TryParse(dr["hours"].ToString(), out hr))
                        retVal[sid] += hr;
                }
            }
            return retVal;
        }


        /// <summary>
        /// 透過日期區間取得獎懲資料,傳入學生ID,開始日期,結束日期,回傳：學生ID,獎懲統計名稱,統計值
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetDisciplineCountByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<string> nameList = new string[] { "大功", "小功", "嘉獎", "大過", "小過", "警告", "留校" }.ToList();

            // 取得獎懲資料
            List<DisciplineRecord> dataList = Discipline.SelectByStudentIDs(StudentIDList);

            foreach (DisciplineRecord data in dataList)
            {
                if (data.OccurDate >= beginDate && data.OccurDate <= endDate)
                {
                    // 初始化
                    if (!retVal.ContainsKey(data.RefStudentID))
                    {
                        retVal.Add(data.RefStudentID, new Dictionary<string, int>());
                        foreach (string str in nameList)
                            retVal[data.RefStudentID].Add(str, 0);
                    }

                    // 獎勵
                    if (data.MeritFlag == "1")
                    {
                        if (data.MeritA.HasValue)
                            retVal[data.RefStudentID]["大功"] += data.MeritA.Value;

                        if (data.MeritB.HasValue)
                            retVal[data.RefStudentID]["小功"] += data.MeritB.Value;

                        if (data.MeritC.HasValue)
                            retVal[data.RefStudentID]["嘉獎"] += data.MeritC.Value;
                    }
                    else if (data.MeritFlag == "0")
                    { // 懲戒
                        if (data.Cleared != "是")
                        {
                            if (data.DemeritA.HasValue)
                                retVal[data.RefStudentID]["大過"] += data.DemeritA.Value;

                            if (data.DemeritB.HasValue)
                                retVal[data.RefStudentID]["小過"] += data.DemeritB.Value;

                            if (data.DemeritC.HasValue)
                                retVal[data.RefStudentID]["警告"] += data.DemeritC.Value;
                        }
                    }
                    else if (data.MeritFlag == "2")
                    {
                        // 留校察看
                        retVal[data.RefStudentID]["留校"]++;
                    }
                }
            }
            return retVal;
        }

        /// <summary>
        /// 透過日期區間取得學生缺曠統計(傳入學生系統編號、開始日期、結束日期；回傳：學生系統編號、獎懲名稱,統計值
        /// </summary>
        /// <param name="StudIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetAttendanceCountByDate(List<StudentRecord> StudRecordList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<PeriodMappingInfo> PeriodMappingList = PeriodMapping.SelectAll();
            // 節次>類別
            Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
            foreach (PeriodMappingInfo rec in PeriodMappingList)
            {
                if (!PeriodMappingDict.ContainsKey(rec.Name))
                    PeriodMappingDict.Add(rec.Name, rec.Type);
            }

            List<AttendanceRecord> attendList = K12.Data.Attendance.SelectByDate(StudRecordList, beginDate, endDate);

            // 計算統計資料
            foreach (AttendanceRecord rec in attendList)
            {
                if (!retVal.ContainsKey(rec.RefStudentID))
                    retVal.Add(rec.RefStudentID, new Dictionary<string, int>());

                foreach (AttendancePeriod per in rec.PeriodDetail)
                {
                    if (!PeriodMappingDict.ContainsKey(per.Period))
                        continue;

                    // ex.一般:曠課
                    //string key = "區間" + PeriodMappingDict[per.Period] + "_" + per.AbsenceType;

                    string key = PeriodMappingDict[per.Period] +  per.AbsenceType;
                    if (!retVal[rec.RefStudentID].ContainsKey(key))
                        retVal[rec.RefStudentID].Add(key, 0);

                    retVal[rec.RefStudentID][key]++;
                }
            }

            return retVal;
        }

        /// <summary>
        /// 透過 ClassID 取得學生為一般，依照年級、班級名稱、座號排序
        /// </summary>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static List<string> GetClassStudentIDList1ByClassID(List<string> ClassIDList)
        {
            List<string> retVal = new List<string>();
            QueryHelper qh = new QueryHelper();
            string query = "select student.id from student inner join class on student.ref_class_id = class.id where student.status=1 and class.id in("+string.Join(",",ClassIDList.ToArray())+") order by class.grade_year,class.class_name,student.seat_no";
            DataTable dt = new DataTable();
            dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
                retVal.Add(dr[0].ToString());

            dt.Clear();
            return retVal;
        }

        
        /// <summary>
        /// 取得系統內學生類別,群組用[]表示,沒有群組直接名稱
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, List<string>> GetStudentTagRefDict()
        {
            // 學生類別,StudentID
            Dictionary<string, List<string>> retVal = new Dictionary<string, List<string>>();
            QueryHelper qh = new QueryHelper();
            string query = "select tag.prefix,tag.name,ref_student_id from tag left join tag_student on tag.id = tag_student.ref_tag_id order by tag.prefix,tag.name";
            DataTable dt = new DataTable();
            dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
            {
                string strP = "",key="",StudID="";

                if(dr["prefix"]!=null)
                    strP = dr["prefix"].ToString();

                if (string.IsNullOrEmpty(strP))
                    key = dr["name"].ToString();
                else
                    key = "[" + strP + "]";

                if (dr["ref_student_id"] != null)
                    StudID = dr["ref_student_id"].ToString();
                
                if (!retVal.ContainsKey(key))
                    retVal.Add(key, new List<string>());

                if (!string.IsNullOrEmpty(StudID))
                    retVal[key].Add(StudID);

            }               
            return retVal;        
        }

        public static Dictionary<string, Dictionary<string,DAO.SubjectDomainName>> GetStudentSCAttendCourse(List<string> StudentIDList,List<string> CourseIDList,string examID)
        {
            Dictionary<string, Dictionary<string, DAO.SubjectDomainName>> retVal = new Dictionary<string, Dictionary<string, DAO.SubjectDomainName>>();
            QueryHelper qh = new QueryHelper();
            string query = "select ref_student_id,course.domain,course.subject,course.credit from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join te_include on course.ref_exam_template_id = te_include.ref_exam_template_id where sc_attend.ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and course.id in(" + string.Join(",", CourseIDList.ToArray()) + ") and te_include.ref_exam_id=" + examID;
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                string id = dr[0].ToString();

                if (!retVal.ContainsKey(id))
                    retVal.Add(id, new Dictionary<string, DAO.SubjectDomainName>());

                string domainName = dr["domain"].ToString();
                string subjectName = dr["subject"].ToString();
                
                if (string.IsNullOrEmpty(domainName))
                    domainName = "彈性課程";

                

                if (!retVal[id].ContainsKey(subjectName))
                {
                    DAO.SubjectDomainName sdn = new DAO.SubjectDomainName();
                    sdn.SubjectName = subjectName;
                    sdn.DomainName = domainName;
                    decimal credit;
                    if (decimal.TryParse(dr["credit"].ToString(), out credit))
                    {
                        sdn.Credit = credit;
                    }
                    
                    retVal[id].Add(subjectName, sdn);
                }
            }
                        
            return retVal;
        }

        /// <summary>
        /// 透過學生系統編號、學年度、學期，取得考試的科目名稱
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, List<string>> GetExamSubjecList(List<string> StudentIDList, int SchoolYear, int Semester)
        {
            Dictionary<string, List<string>> retVal = new Dictionary<string, List<string>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select distinct ref_exam_id,course.subject from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join te_include on course.ref_exam_template_id = te_include.ref_exam_template_id where sc_attend.ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and course.school_year=" + SchoolYear + " and  course.semester=" + Semester + " and course.subject is not null order by ref_exam_id,subject";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr[0].ToString();

                    if (!retVal.ContainsKey(id))
                        retVal.Add(id, new List<string>());

                    string subjectName = dr["subject"].ToString();

                    if (!retVal[id].Contains(subjectName))
                        retVal[id].Add(subjectName);
                }
            }
            return retVal;
        }

         /// <summary>
        /// 日常生活表現名稱對照使用
        /// </summary>
        internal static Dictionary<string, string> DLBehaviorConfigNameDict = new Dictionary<string, string>();

        /// <summary>
        /// 日常生活表現子項目名稱,呼叫GetDLBehaviorConfigNameDict 一同取得
        /// </summary>
        internal static Dictionary<string, List<string>> DLBehaviorConfigItemNameDict = new Dictionary<string, List<string>>();

        /// <summary>
        /// XML 內解析子項目名稱
        /// </summary>
        /// <param name="elm"></param>
        /// <returns></returns>
        internal static List<string> ParseItems(XElement elm)
        {
            List<string> retVal = new List<string>();

            foreach (XElement subElm in elm.Elements("Item"))
            {
                // 因為社團功能，所以要將"社團活動" 字不放入
                string name = subElm.Attribute("Name").Value;
                if(name !="社團活動")
                    retVal.Add(name);
            }
            return retVal;
        }

        /// <summary>
        /// 取得日常生活表現設定名稱
        /// </summary>
        /// <returns></returns>
        internal static Dictionary<string, string> GetDLBehaviorConfigNameDict()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            try
            {
                DLBehaviorConfigItemNameDict.Clear();
                // 包含新竹
                K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];
                if (!string.IsNullOrEmpty(cd["DailyBehavior"]))
                {
                    string key = "日常行為表現";
                    //日常行為表現
                    XElement e1 = XElement.Parse(cd["DailyBehavior"]);
                    string name = e1.Attribute("Name").Value;
                    retVal.Add(key, name);

                    // 日常生活表現子項目
                    List<string> items = ParseItems(e1);
                    if (items.Count > 0)
                        DLBehaviorConfigItemNameDict.Add(key, items);

                }
                if (!string.IsNullOrEmpty(cd["OtherRecommend"]))
                {
                    //其它表現
                    XElement e2 = XElement.Parse(cd["OtherRecommend"]);
                    string name = e2.Attribute("Name").Value;
                    retVal.Add("其它表現", name);
                }
                if (!string.IsNullOrEmpty(cd["DailyLifeRecommend"]))
                {
                    //日常生活表現具體建議
                    XElement e3 = XElement.Parse(cd["DailyLifeRecommend"]);
                    string name = e3.Attribute("Name").Value;
                    retVal.Add("綜合評語", name);
                }
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("日常生活表現設定檔解析失敗!" + ex.Message);
            }

            return retVal;
        }

        /// <summary>
        /// 取得學生日常生活表現
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, DAO.StudTextScoreXML> GetStudentTextScoreDict(List<string> StudentIDList, int SchoolYear, int Semester)
        {
            Dictionary<string, DAO.StudTextScoreXML> retVal = new Dictionary<string, DAO.StudTextScoreXML>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,text_score from sems_moral_score where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and school_year=" + SchoolYear + " and semester=" + Semester;
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["ref_student_id"].ToString();

                    if(!retVal.ContainsKey(sid))
                        retVal.Add(sid, new DAO.StudTextScoreXML(dr["text_score"].ToString()));
                }
            }
            return retVal;
        }

    }
}
