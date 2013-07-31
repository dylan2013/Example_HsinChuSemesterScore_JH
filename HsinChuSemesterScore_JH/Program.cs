using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Presentation;
using FISCA.Permission;

namespace HsinChuSemesterScore_JH
{
    /// <summary>
    /// 新竹國中學期成績(通知單與證明單)
    /// </summary>
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["學生", "資料統計"];
            rbItem1["報表"]["成績相關報表"]["學期成績單(測試版)"].Enable = UserAcl.Current["JH.Student.HsinChuSemesterScore_JH_Student"].Executable;
            rbItem1["報表"]["成績相關報表"]["學期成績單(測試版)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    PrintForm pf = new PrintForm(K12.Presentation.NLDPanels.Student.SelectedSource);
                    pf.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選學生");
                    return;
                }
            };

            RibbonBarItem rbItem2 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbItem2["報表"]["成績相關報表"]["學期成績單(測試版)"].Enable = UserAcl.Current["JH.Student.HsinChuSemesterScore_JH_Class"].Executable;
            rbItem2["報表"]["成績相關報表"]["學期成績單(測試版)"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {
                    List<string> StudentIDList = Utility.GetClassStudentIDList1ByClassID(K12.Presentation.NLDPanels.Class.SelectedSource);
                    PrintForm pf = new PrintForm(StudentIDList);
                    pf.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選班級");
                    return;
                }

            };
            // 學期成績單
            Catalog catalog1a = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog1a.Add(new RibbonFeature("JH.Student.HsinChuSemesterScore_JH_Student", "學期成績單(測試版)"));

            // 學期成績單
            Catalog catalog1b = RoleAclSource.Instance["班級"]["功能按鈕"];
            catalog1b.Add(new RibbonFeature("JH.Student.HsinChuSemesterScore_JH_Class", "學期成績單(測試版)"));
        }
    }
}
