using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace HsinChuSemesterScore_JH.DAO
{
    /// <summary>
    /// 學生學期成績
    /// </summary>
    public class StudentScore
    {
        /// <summary>
        /// 學生編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 領域成績，key領域名稱
        /// </summary>
        public Dictionary<string, DomainScore> DomainScoreDict = new Dictionary<string, DomainScore>();

        /// <summary>
        /// 科目成績，key領域名稱，主要用於各科目成績依領域分開
        /// </summary>
        public Dictionary<string, List<SubjectScore>> SubjectScoreDict = new Dictionary<string, List<SubjectScore>>();

        /// <summary>
        /// 領域文字描述
        /// </summary>
        public List<string> DomainTextList = new List<string>();

        /// <summary>
        /// 科目文字描述
        /// </summary>
        public List<string> SubjecTextList = new List<string>();

        /// <summary>
        /// 課程學習領域成績
        /// </summary>
        public decimal? CourseLearnScore { get; set; }

        /// <summary>
        /// 學習領域成績
        /// </summary>
        public decimal? LearnDomainScore { get; set; }
    }
}
