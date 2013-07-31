using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace HsinChuSemesterScore_JH.DAO
{
    /// <summary>
    /// 學生日常生活表現文字評量
    /// </summary>
    public class StudTextScoreXML
    {
        public string StudentID { get; set; }

        private XElement _DataXML;

        public StudTextScoreXML(string strXML)
        {
            if (!string.IsNullOrEmpty(strXML))
            {
                try
                {
                    _DataXML = XElement.Parse(strXML);
                }
                catch(Exception ex)
                {}
            }

            if (_DataXML == null)
                _DataXML = new XElement("Null");        
        }

        public XElement GetAllXML()
        {
            return _DataXML;
        }

        /// <summary>
        /// 日常行為表現
        /// </summary>
        /// <param name="GroupName"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public string GetDailyBehavior(string GroupName, string ItemName)
        {
            string retVal = "";
            if (_DataXML.Element("DailyBehavior") != null)
            {
                //if (_DataXML.Element("DailyBehavior").Attribute("Name").Value == GroupName)
                //{
                    foreach(XElement itemElm in _DataXML.Element("DailyBehavior").Elements("Item"))
                    {
                        if (itemElm.Attribute("Name").Value == ItemName)
                        {
                            retVal = itemElm.Attribute("Degree").Value;
                            break;
                        }
                    }
                //}
            }

            return retVal;
        }

        /// <summary>
        /// 其他表現
        /// </summary>
        /// <param name="Name"></param>
        /// <returns></returns>
        public string GetOtherRecommend(string Name)
        {
            string retVal = "";
            if (_DataXML.Element("OtherRecommend") != null)
                //if (_DataXML.Element("OtherRecommend").Attribute("Name").Value == Name)
                    retVal = _DataXML.Element("OtherRecommend").Attribute("Description").Value;

            return retVal;
        }

        /// <summary>
        /// 綜合表現
        /// </summary>
        /// <param name="Name"></param>
        /// <returns></returns>
        public string GetDailyLifeRecommend(string Name)
        {
            string retVal = "";

            if (_DataXML.Element("DailyLifeRecommend") != null)
                //if (_DataXML.Element("DailyLifeRecommend").Attribute("Name").Value == Name)
                    retVal = _DataXML.Element("DailyLifeRecommend").Attribute("Description").Value;

            return retVal;
        }
    }
}
