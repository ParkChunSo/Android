using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JSON_ExcelDirectionalConverter
{
    public enum convertingMode { WJSONToWExcel, WExcelToWJSON, CJSONToCExcel, CExcelToEJSON };
    public partial class Form1 : Form
    {
        convertingMode currentCuonvertingMode;
        List<string> filePathList;

        public Form1()
        {
            InitializeComponent();
            cb_modeSelect.DropDownStyle = ComboBoxStyle.DropDownList;

            btn_addFiles.Enabled = false;
            btn_clearList.Enabled = false;
            btn_convert.Enabled = false;
            btn_removeFiles.Enabled = false;
            tbx_fileNmae.Enabled = false;

            filePathList = new List<string>();
        }

        private void cb_modeSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            btn_addFiles.Enabled = true;
            btn_clearList.Enabled = true;
            btn_convert.Enabled = true;
            btn_removeFiles.Enabled = true;
            tbx_fileNmae.Enabled = true;

            lb_fileList.Items.Clear();
            filePathList.Clear();

            switch (cb_modeSelect.SelectedIndex)
            {
                case 0:
                    currentCuonvertingMode = convertingMode.WJSONToWExcel;
                    break;
                case 1:
                    currentCuonvertingMode = convertingMode.WExcelToWJSON;
                    break;
                case 2:
                    currentCuonvertingMode = convertingMode.CJSONToCExcel;
                    break;
                case 3:
                    currentCuonvertingMode = convertingMode.CExcelToEJSON;
                    break;
                default:
                    MessageBox.Show("불가능한 모드값 선택됨. 프로그램 종료");
                    this.Close();
                    return;
            }
        }

        private void btn_addFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;

            if (currentCuonvertingMode == convertingMode.WJSONToWExcel || currentCuonvertingMode == convertingMode.CJSONToCExcel)
                openFileDialog.Filter = "JSON Files|*.json;";
            else if (currentCuonvertingMode == convertingMode.WExcelToWJSON || currentCuonvertingMode == convertingMode.CExcelToEJSON)
                openFileDialog.Filter = "Excel Files|*.xlsx;";
            else { MessageBox.Show("불가능한 모드값 선택됨. 프로그램 종료"); this.Close(); return; }
            

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePathList.AddRange(openFileDialog.FileNames);
                for (int i = 0; i < openFileDialog.FileNames.Length; i++)
                {
                    string tempFileName = openFileDialog.FileNames[i].Substring(openFileDialog.FileNames[i].LastIndexOf("\\") + 1);
                    lb_fileList.Items.Add(tempFileName);
                }

            }
        }

        private void btn_removeFiles_Click(object sender, EventArgs e)
        {
            ListBox.SelectedIndexCollection lbSIC = new ListBox.SelectedIndexCollection(lb_fileList);
            lbSIC = lb_fileList.SelectedIndices;
            int roofCount = lbSIC.Count;
            for (int i = 0; i < roofCount; i++)
            {
                int inextNum = lbSIC[0];
                filePathList.RemoveAt(inextNum);
                lb_fileList.Items.RemoveAt(inextNum);
            }
        }

        private void btn_clearList_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("모든 파일 목록을 초기화하시겠습니까?", "초기화", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                lb_fileList.Items.Clear();
                filePathList.Clear();
            }
        }

        private void btn_convert_Click(object sender, EventArgs e)
        {
            if(currentCuonvertingMode == convertingMode.CExcelToEJSON || currentCuonvertingMode == convertingMode.CJSONToCExcel)
            {

                Cursor.Current = Cursors.WaitCursor;

                CrossJEConverter_v2 jeConverter = new CrossJEConverter_v2(currentCuonvertingMode, tbx_fileNmae.Text);
                MessageBox.Show(jeConverter.convertFiles(filePathList));

                Cursor.Current = Cursors.Default;
            }
            else
            {
                Cursor.Current = Cursors.WaitCursor;

                WorkJEConverter jeConverter = new WorkJEConverter(currentCuonvertingMode);
                MessageBox.Show(jeConverter.convertFiles(filePathList));

                Cursor.Current = Cursors.Default;
            }
            
        }

        
    }
}

