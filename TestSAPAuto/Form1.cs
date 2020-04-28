using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAPAuto;

namespace TestSAPAuto {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.login("NRP", "100", "28066351", "Hackmeifyoucan@123", "ZH");
        }

        private void button2_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.TCODEInput("ZFCN130");
        }

        private void button3_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            //auto.TextInput(txtID.Text, txtValue.Text);
        }

        private void button4_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            //auto.TextInput(txtID.Text, txtValue.Text);
        }

        private void button5_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.TextInputByID(txtID.Text, txtValue.Text);
        }

        private void button6_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.ButtonPressByID_Using_MainWindowName("客户对帐单", "wnd[2]/tbar[0]/btn[0]");
            //auto.ButtonPressByID(txtID.Text);
        }

        private void button7_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.ContextMenuSelectItem("", "Spreadsheet...");
        }

        private void button8_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.LogOnButtonClick();
        }

        private void button9_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.TableRow_Select("wnd[1]/usr/tblSAPLAQ_INT_FUNCTIONSTCH_FUNCAREAS", "6");
        }

        private void button10_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.MenuSelectionByID("wnd[0]/mbar/menu[4]/menu[8]");
        }

        private void button11_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.SaveAsExcelWindow("Customer Statement for Sales Report", @"D:\Users\28066351\Documents\Projects\RPA Metabots\SAPAuto\TestSave");
        }

        private void button12_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.TreeView_Select("wnd[0]/shellcont/shell/shellcont[0]/shell", "000006");
        }

        private void btnGetWinText_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            MessageBox.Show(auto.GetWindowText("wnd[1]"));            
        }

        private void btnClickUsingUIAutomatio_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.ButtonPress_UsingUIAutomation("SAP 轻松访问", "", "Close");
        }

        private void ctnGetLabel_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            var output =  auto.GetLabelText("作业总览", "", "", "wnd[0]/usr/lbl[64,3]");
            MessageBox.Show("label text:" + output);
        }

        private void btnUpdate_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
           System.Windows.Forms.MessageBox.Show(auto.CheckAAError_IfError_Close());
        }

        private void button13_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
           // auto.testbusyornot();
        }

        private void button14_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            System.Windows.Forms.MessageBox.Show(auto.IsWindow_Exists_UsingUIAutomation_ThreeWindows("总帐科目行项目显示 总帐视图", "总帐科目 的多种选择", "总帐科目 的多种选择").ToString());
        }

        private void button15_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.Grid_SelectCell_AndDoubleClick("wnd[0]/usr/cntlGRID1/shellcont/shell", "0", "EBELN");//MESSTXT1
        }

        private void button16_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.ComboBox_SelectItem_Using_Key("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST", "   3");
        }

        private void button17_Click(object sender, EventArgs e) {
            SAPAuto1 auto = new SAPAuto1();
            auto.ComboBox_SelectItem_Using_ItemTextcontains("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB1:SAPLMEGUI:6000/cmbDYN_6000-LIST", "25027");

        }
    }
}
