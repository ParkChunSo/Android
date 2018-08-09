using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using JSON_ExcelDirectionalConverter.WtagClasses;

namespace JSON_ExcelDirectionalConverter
{
    public class WorkJEConverter
    {
        const string STR_CONVERTING_SUCCESS = "SUCCESS";

        private convertingMode m_currentConvertingMode;
        private string m_path;
        private string m_savePath;

        private bool m_EtoJNullRemoveCheck = false;
        
        private static string[] sheet1ColHeader = {"number", "id", "confuseQt","confuseQf","confuseLat","confuseSat","question",
                                                     "question_en","question_tagged", "questionType", "questionFocus", "questionSAT",
                                                     "questionLAT", "etriQtCheck","etriQfCheck","etriLatCheck","etriSatCheck","etriQt","etriQf","etriLat","etriSat", "ans_text1", "ans_text_en1", "ans_text_tagged1",
                                                     "ans_text_syn1", "ans_start1", "ans_end1", "ans_text2", "ans_text_en2", "ans_text_tagged2",
                                                     "ans_text_syn2", "ans_start2", "ans_end2","ans_text3", "ans_text_en3", "ans_text_tagged3",
                                                     "ans_text_syn3", "ans_start3", "ans_end3", "paragraphs_num"};

        private int sheet1ColCount = sheet1ColHeader.Length;
        private int sheet1RowCount;

        private static string[] sheet2ColHeader = {"number", "version", "creator", "progress", "formatt", "time",
                                                      "title",  "context", "context_en", 
                                                     "context_tagged"};
        private int sheet2ColCount = sheet2ColHeader.Length;
        private int sheet2RowCount;

        private Excel.Application objApp;
        private Excel.Workbooks objWorkbooks;
        private Excel.Workbook objWorkbook;
        private Excel.Sheets objWorksheets;
        private Excel.Worksheet objWorksheet;
        private Excel.Range range;

        private Excel.XlHAlign HCENTER = Excel.XlHAlign.xlHAlignCenter;

        public WorkJEConverter(convertingMode mode)
        {
            m_currentConvertingMode = mode;
            sheet1RowCount = 0;
            sheet2RowCount = 0;
        }

        public string convertFiles(IList<string> filePaths)
        {
            foreach (var item in filePaths)
            {
                string tempStat = convert(item);
                if (tempStat != STR_CONVERTING_SUCCESS)
                {
                    return tempStat;
                }
            }

            return "모든 파일 변환 성공";
        }

        private string convert(string filePath)
        {
            m_path = filePath;

            var missing = Type.Missing;

            objApp = new Excel.Application();
            objWorkbooks = objApp.Workbooks;

            TopName topName;
            Data[] data;
            Paragraphs[][] paragraphs;
            Qas[][][] qas;
            Answers[][][][] answers;

            object[,] sheet1ValueArray;
            object[,] sheet2ValueArray;

            int countParagraphs = 0;
            int countQas = 0;
            int currentRow = 0;

            bool excelOpen = false;

            try
            {
               
                if (m_currentConvertingMode == convertingMode.WJSONToWExcel)
                {
                    #region JSON -> Excel 변환
                    
                        // ** name1 영역 파싱
                        topName = JsonConvert.DeserializeObject<TopName>(File.ReadAllText(m_path));

                        // name2 영역 파싱
                        data = new Data[topName.data.Count];
                        for (int i = 0; i < data.Length; i++)
                        {
                            data[i] = JsonConvert.DeserializeObject<Data>(topName.data[i].ToString());
                        }

                        // ** name3 영역 파싱
                        paragraphs = new Paragraphs[data.Length][];
                        for (int i = 0; i < data.Length; i++)
                        {
                            paragraphs[i] = new Paragraphs[data[i].paragraphs.Count];
                            for (int j = 0; j < data[i].paragraphs.Count; j++)
                            {
                                paragraphs[i][j] = JsonConvert.DeserializeObject<Paragraphs>(data[i].paragraphs[j].ToString());
                                countParagraphs++;
                            }
                        }

                        // ** name4 영역 파싱
                        qas = new Qas[data.Length][][];
                        for (int i = 0; i < data.Length; i++)
                        {
                            qas[i] = new Qas[paragraphs[i].Length][];
                            for (int j = 0; j < paragraphs[i].Length; j++)
                            {
                                qas[i][j] = new Qas[paragraphs[i][j].qas.Count];
                                for (int k = 0; k < paragraphs[i][j].qas.Count; k++)
                                {
                                    qas[i][j][k] = JsonConvert.DeserializeObject<Qas>(paragraphs[i][j].qas[k].ToString());
                                    countQas++;
                                }
                            }
                        }

                        // ** name5 영역 파싱
                        answers = new Answers[data.Length][][][];
                        for (int i = 0; i < data.Length; i++)
                        {
                            answers[i] = new Answers[paragraphs[i].Length][][];
                            for (int j = 0; j < paragraphs[i].Length; j++)
                            {
                                answers[i][j] = new Answers[qas[i][j].Length][];
                                for (int k = 0; k < qas[i][j].Length; k++)
                                {
                                    answers[i][j][k] = new Answers[qas[i][j][k].answers.Count];
                                    for (int m = 0; m < qas[i][j][k].answers.Count; m++)
                                    {
                                        answers[i][j][k][m] = JsonConvert.DeserializeObject<Answers>(qas[i][j][k].answers[m].ToString());
                                    }
                                }
                            }
                        }
                   

                    //// ** ':' 문자가 데이터 내에 포함되어 있는지 검사
                    //// * name1
                    //if (topName.formatt != null && topName.formatt.Contains(':'))
                    //{
                    //    return "name1-formatt 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //}
                    //if (topName.version != null && topName.version.Contains(':'))
                    //{
                    //    return "name1-version 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //}
                    //if (topName.creator != null && topName.creator.Contains(':'))
                    //{
                    //    return "name1-creator 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //}

                    //// * name2
                    //foreach (var item in data)
                    //{
                    //    if (item.title != null && item.title.Contains(':'))
                    //    {
                    //        return "name2-title 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //    }
                    //}

                    //// * name3
                    //foreach (var arr in paragraphs)
                    //{
                    //    foreach (var item in arr)
                    //    {
                    //        if (item.context != null && item.context.Contains(':'))
                    //        {
                    //            return "name3-context 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //        }
                    //        if (item.context_original != null && item.context_original.Contains(':'))
                    //        {
                    //            return "name3-context_original 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //        }
                    //        if (item.context_en != null && item.context_en.Contains(':'))
                    //        {
                    //            return "name3-context_en 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //        }
                    //        if (item.context_tagged != null)
                    //        {
                    //            foreach (var item2 in item.context_tagged)
                    //            {
                    //                if (item2 != null && item2.Contains(':'))
                    //                {
                    //                    return "name3-context_tagged 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //                }
                    //            }
                    //        }
                    //    }
                    //}

                    //// * name4
                    //foreach (var arr1 in qas)
                    //{
                    //    foreach (var arr2 in arr1)
                    //    {
                    //        foreach (var item in arr2)
                    //        {
                    //            if (item.confuse != null && item.confuse.Contains(':'))
                    //            {
                    //                return "name4-confuse 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.id != null && item.id.Contains(':'))
                    //            {
                    //                return "name4-id 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.question != null && item.question.Contains(':'))
                    //            {
                    //                return "name4-question 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.question_original != null && item.question_original.Contains(':'))
                    //            {
                    //                return "name4-question_original 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.question_en != null && item.question_en.Contains(':'))
                    //            {
                    //                return "name4-question_en 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.questionType != null && item.questionType.Contains(':'))
                    //            {
                    //                return "name4-questionType 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.questionFocus != null && item.questionFocus.Contains(':'))
                    //            {
                    //                return "name4-questionFocus 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.questionSAT != null && item.questionSAT.Contains(':'))
                    //            {
                    //                return "name4-questionSAT 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.questionLAT != null && item.questionLAT.Contains(':'))
                    //            {
                    //                return "name4-questionLAT 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //            }
                    //            if (item.question_tagged != null)
                    //            {
                    //                foreach (var item2 in item.question_tagged)
                    //                {
                    //                    if (item2 != null && item2.Contains(':'))
                    //                    {
                    //                        return "name4-question_tagged 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //                    }
                    //                }
                    //            }
                    //        }
                    //    }
                    //}

                    //// * name5
                    //foreach (var arr1 in answers)
                    //{
                    //    foreach (var arr2 in arr1)
                    //    {
                    //        foreach (var arr3 in arr2)
                    //        {
                    //            foreach (var item in arr3)
                    //            {
                    //                if (item.text != null && item.text.Contains(':'))
                    //                {
                    //                    return "name5-text 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //                }
                    //                if (item.text_original != null && item.text_original.Contains(':'))
                    //                {
                    //                    return "name5-text_original 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //                }
                    //                if (item.text_en != null && item.text_en.Contains(':'))
                    //                {
                    //                    return "name5-text_en 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //                }
                    //                if (item.text_tagged != null)
                    //                {
                    //                    foreach (var item2 in item.text_tagged)
                    //                    {
                    //                        if (item2 != null && item2.Contains(':'))
                    //                        {
                    //                            return "name5-text_tagged 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //                        }
                    //                    }
                    //                }
                    //                if (item.text_syn != null)
                    //                {
                    //                    foreach (var item2 in item.text_syn)
                    //                    {
                    //                        if (item2 != null && item2.Contains(':'))
                    //                        {
                    //                            return "name5-text_syn 영역에 구분자(:) 발견.\r\n파일: " + filePath;
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }
                    //    }
                    //}

                    // ** sheet1ValueArray & sheet2ValueArray 영역 크기 지정
                    sheet1RowCount = countQas;
                    sheet2RowCount = countParagraphs;

                    sheet1ValueArray = new object[sheet1RowCount, sheet1ColCount];
                    sheet2ValueArray = new object[sheet2RowCount, sheet2ColCount];

                    // ** sheet1ValueArray & sheet2ValueArray에 데이터 입력
                    // * paragraph 순번 & name1 영역
                    for (int row = 0; row < sheet2RowCount; row++)
                    {
                        sheet2ValueArray[row, 0] = row + 1;
                        sheet2ValueArray[row, 1] = topName.version;
                        sheet2ValueArray[row, 2] = topName.creator;
                        sheet2ValueArray[row, 3] = topName.progress;
                        sheet2ValueArray[row, 4] = topName.formatt;
                        sheet2ValueArray[row, 5] = topName.time;
                        //sheet2ValueArray[row, 0] = row + 1;
                        //sheet2ValueArray[row, 1] = topName.time;
                        //sheet2ValueArray[row, 2] = topName.formatt;
                        //sheet2ValueArray[row, 3] = topName.progress;
                        //sheet2ValueArray[row, 4] = topName.version;
                        //sheet2ValueArray[row, 5] = topName.creator;
                    }

                    // * name2 & name3 영역
                    currentRow = 0;
                    for (int d = 0; d < data.Length; d++)
                    {
                        for (int p = 0; p < paragraphs[d].Length; p++)
                        {
                            sheet2ValueArray[currentRow, 6] = data[d].title;
                            sheet2ValueArray[currentRow, 7] = paragraphs[d][p].context;
                            sheet2ValueArray[currentRow, 8] = paragraphs[d][p].context_en;
                            sheet2ValueArray[currentRow, 9] = paragraphs[d][p].context_tagged;

                            //if (paragraphs[d][p].context_tagged == null)
                            //{
                            //    sheet2ValueArray[currentRow, 9] = null;
                            //    currentRow++;
                            //    continue;
                            //}
                            //string tempStr = "";
                            //for (int ct = 0; ct < paragraphs[d][p].context_tagged.Count; ct++)
                            //{
                            //    if (tempStr == "")
                            //    {
                            //        tempStr = paragraphs[d][p].context_tagged[ct];
                            //    }
                            //    else
                            //    {
                            //        if (paragraphs[d][p].context_tagged[ct] != null)
                            //        {
                            //            tempStr = tempStr + ":" + paragraphs[d][p].context_tagged[ct];
                            //        }
                            //    }
                            //    if (tempStr == null)
                            //        tempStr = "";
                            //}
                            //sheet2ValueArray[currentRow, 10] = (tempStr == "") ? null : tempStr;

                            currentRow++;
                        }
                    }

                    // * name4 영역
                    currentRow = 0;
                    int currentParaNum = 1;
                    for (int d = 0; d < data.Length; d++)
                    {
                        for (int p = 0; p < paragraphs[d].Length; p++)
                        {
                            for (int q = 0; q < qas[d][p].Length; q++)
                            {
                                sheet1ValueArray[currentRow, 0] = currentRow + 1;
                                sheet1ValueArray[currentRow, 1] = qas[d][p][q].id;
                                sheet1ValueArray[currentRow, 2] = qas[d][p][q].confuseQt;
                                sheet1ValueArray[currentRow, 3] = qas[d][p][q].confuseQf;
                                sheet1ValueArray[currentRow, 4] = qas[d][p][q].confuseLat;
                                sheet1ValueArray[currentRow, 5] = qas[d][p][q].confuseSat;
                                sheet1ValueArray[currentRow, 6] = qas[d][p][q].question;
                                //sheet1ValueArray[currentRow, 7] = qas[d][p][q].question_original;
                                sheet1ValueArray[currentRow, 7] = qas[d][p][q].question_en;
                                sheet1ValueArray[currentRow, 8] = qas[d][p][q].question_tagged;
                                sheet1ValueArray[currentRow, 9] = qas[d][p][q].questionType;
                                sheet1ValueArray[currentRow, 10] = qas[d][p][q].questionFocus;
                                sheet1ValueArray[currentRow, 11] = qas[d][p][q].questionSAT;
                                sheet1ValueArray[currentRow, 12] = qas[d][p][q].questionLAT;
                                sheet1ValueArray[currentRow, 13] = qas[d][p][q].etriQtCheck;
                                sheet1ValueArray[currentRow, 14] = qas[d][p][q].etriQfCheck;
                                sheet1ValueArray[currentRow, 15] = qas[d][p][q].etriLatCheck;
                                sheet1ValueArray[currentRow, 16] = qas[d][p][q].etriSatCheck;
                                sheet1ValueArray[currentRow, 17] = qas[d][p][q].etriQt;//
                                sheet1ValueArray[currentRow, 18] = qas[d][p][q].etriQf;//
                                sheet1ValueArray[currentRow, 19] = qas[d][p][q].etriLat;//
                                sheet1ValueArray[currentRow, 20] = qas[d][p][q].etriSat;//

                                sheet1ValueArray[currentRow, 39] = currentParaNum;

                                //if (qas[d][p][q].question_tagged == null)
                                //{
                                //    sheet1ValueArray[currentRow, 9] = null;
                                //    currentRow++;
                                //    continue;
                                //}
                                //string tempStr = "";
                                //for (int qt = 0; qt < qas[d][p][q].question_tagged.Count; qt++)
                                //{
                                //    if (tempStr == "")
                                //    {
                                //        tempStr = qas[d][p][q].question_tagged[qt];
                                //    }
                                //    else
                                //    {
                                //        if (qas[d][p][q].question_tagged[qt] != null)
                                //        {
                                //            tempStr = tempStr + ":" + qas[d][p][q].question_tagged[qt];
                                //        }
                                //    }
                                //    if (tempStr == null)
                                //        tempStr = "";
                                //}
                                //sheet1ValueArray[currentRow, 7] = (tempStr == "") ? null : tempStr;
                                //currentRow++;
                                currentRow++;
                            }
                           
                            currentParaNum++;
                        }
                    }

                    // * name5 영역
                    currentRow = 0;
                    for (int d = 0; d < data.Length; d++)
                    {
                        for (int p = 0; p < paragraphs[d].Length; p++)
                        {
                            for (int q = 0; q < qas[d][p].Length; q++)
                            {
                                if (qas[d][p][q].answers.Count > 3)
                                {
                                    return "정답의 개수가 3개 초과인 문제가 있습니다.\r\n파일: " + filePath;
                                }

                                int answerStartColNum = 21;
                                for (int a = 0; a < answers[d][p][q].Length; a++)
                                {
                                    sheet1ValueArray[currentRow, answerStartColNum] = answers[d][p][q][a].text;
                                    //sheet1ValueArray[currentRow, answerStartColNum + 1] = answers[d][p][q][a].text_original;
                                    sheet1ValueArray[currentRow, answerStartColNum + 1] = answers[d][p][q][a].text_en;
                                    sheet1ValueArray[currentRow, answerStartColNum + 2] = answers[d][p][q][a].text_tagged;
                                    sheet1ValueArray[currentRow, answerStartColNum + 3] = answers[d][p][q][a].text_syn;
                                    sheet1ValueArray[currentRow, answerStartColNum + 4] = answers[d][p][q][a].answer_start;
                                    sheet1ValueArray[currentRow, answerStartColNum + 5] = answers[d][p][q][a].answer_end;
                                    //string tempStr = "";
                                    //if (answers[d][p][q][a].text_tagged == null)
                                    //{
                                    //    sheet1ValueArray[currentRow, answerStartColNum + 3] = null;
                                    //}
                                    //else
                                    //{
                                    //    for (int tt = 0; tt < answers[d][p][q][a].text_tagged.Count; tt++)
                                    //    {
                                    //        if (tempStr == "")
                                    //        {
                                    //            tempStr = answers[d][p][q][a].text_tagged[tt];
                                    //        }
                                    //        else
                                    //        {
                                    //            if (answers[d][p][q][a].text_tagged[tt] != null)
                                    //            {
                                    //                tempStr = tempStr + ":" + answers[d][p][q][a].text_tagged[tt];
                                    //            }
                                    //        }
                                    //        if (tempStr == null)
                                    //            tempStr = "";
                                    //    }
                                    //    sheet1ValueArray[currentRow, answerStartColNum + 2] = tempStr;
                                    //}
                                    //tempStr = "";
                                    //if (answers[d][p][q][a].text_syn == null)
                                    //{
                                    //    sheet1ValueArray[currentRow, answerStartColNum + 3] = null;
                                    //}
                                    //else
                                    //{
                                    //    for (int ts = 0; ts < answers[d][p][q][a].text_syn.Count; ts++)
                                    //    {
                                    //        if (tempStr == "")
                                    //        {
                                    //            tempStr = answers[d][p][q][a].text_syn[ts];
                                    //        }
                                    //        else
                                    //        {
                                    //            if (answers[d][p][q][a].text_syn[ts] != null)
                                    //            {
                                    //                tempStr = tempStr + ":" + answers[d][p][q][a].text_syn[ts];
                                    //            }
                                    //        }
                                    //        if (tempStr == null)
                                    //            tempStr = "";
                                    //    }
                                    //    sheet1ValueArray[currentRow, answerStartColNum + 3] = tempStr;
                                    //}
                                    answerStartColNum += 6;
                                }
                                currentRow++;
                            }
                        }
                    }

                    // ** 엑셀로 출력
                    excelOpen = true;
                    objWorkbook = objWorkbooks.Add(missing);
                    objWorksheets = objWorkbook.Worksheets;

                    // * sheet2 부분 적용
                    objWorksheet = (Excel.Worksheet)objWorksheets.get_Item(1);
                    objWorksheet.Name = "Paragraphs";

                    range = objWorksheet.get_Range("A1", "J1");
                    range.HorizontalAlignment = HCENTER;
                    range.Interior.Color = Color.FromArgb(142, 169, 219);
                    range.Value2 = sheet2ColHeader;
                    Marshal.ReleaseComObject(range);

                    Excel.Range c1 = objWorksheet.Cells[2, 1];
                    Excel.Range c2 = objWorksheet.Cells[sheet2RowCount + 1, sheet2ColCount];
                    range = objWorksheet.get_Range(c1, c2);
                    range.Value = sheet2ValueArray;
                    Marshal.FinalReleaseComObject(c1);
                    Marshal.FinalReleaseComObject(c2);
                    Marshal.FinalReleaseComObject(range);

                    Marshal.ReleaseComObject(objWorksheet);

                    // * sheet1 부분 적용
                    objWorksheet = (Excel.Worksheet)objWorksheets.Add(missing, missing, missing, missing);
                    objWorksheet.Name = "Qas";

                    range = objWorksheet.get_Range("A1", "AN1");
                    range.HorizontalAlignment = HCENTER;
                    range.Interior.Color = Color.FromArgb(142, 169, 219);
                    range.Value2 = sheet1ColHeader;
                    Marshal.ReleaseComObject(range);

                    c1 = objWorksheet.Cells[2, 1];
                    c2 = objWorksheet.Cells[sheet1RowCount + 1, sheet1ColCount];
                    range = objWorksheet.get_Range(c1, c2);
                    range.Value = sheet1ValueArray;
                    Marshal.FinalReleaseComObject(c1);
                    Marshal.FinalReleaseComObject(c2);
                    Marshal.FinalReleaseComObject(range);

                    Marshal.FinalReleaseComObject(objWorksheet);
                    Marshal.FinalReleaseComObject(objWorksheets);


                    m_savePath = Path.ChangeExtension(m_path, "xlsx");
                    FileInfo fi = new FileInfo(m_savePath);
                    if (fi.Exists)
                    {
                        fi.Delete();
                    }

                    objWorkbook.SaveAs(m_savePath, Excel.XlFileFormat.xlOpenXMLWorkbook,
                    missing, missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    Excel.XlSaveConflictResolution.xlUserResolution, true, missing, missing, missing);

                    objWorkbook.Close(false, missing, missing);
                    objWorkbooks.Close();
                    objApp.Quit();

                    Marshal.FinalReleaseComObject(objWorkbook);
                    Marshal.FinalReleaseComObject(objWorkbooks);
                    Marshal.FinalReleaseComObject(objApp);

                    objApp = null;
                    excelOpen = false;
                    #endregion
                }
                else
                {
                    #region Excel -> JSON 변환

                    // ** Excel 파일 불러와서 object 이중배열에 데이터 입력
                    excelOpen = true;
                    objWorkbook = objWorkbooks.Open(m_path);
                    objWorksheets = objWorkbook.Worksheets;

                    objWorksheet = (Excel.Worksheet)objWorksheets[1];
                    range = objWorksheet.UsedRange;
                    sheet1ValueArray = (object[,])range.get_Value(missing);
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(objWorksheet);

                    objWorksheet = (Excel.Worksheet)objWorksheets[2];
                    range = objWorksheet.UsedRange;
                    sheet2ValueArray = (object[,])range.get_Value(missing);
                    Marshal.FinalReleaseComObject(range);
                    Marshal.FinalReleaseComObject(objWorksheet);

                    Marshal.FinalReleaseComObject(objWorksheets);

                    objWorkbook.Close(false, missing, missing);
                    objWorkbooks.Close();
                    objApp.Quit();

                    Marshal.FinalReleaseComObject(objWorkbook);
                    Marshal.FinalReleaseComObject(objWorkbooks);
                    Marshal.FinalReleaseComObject(objApp);

                    objApp = null;
                    excelOpen = false;

                    // ** sheet1, sheet2 object 이중배열의 데이터를 JSON 태그 클래스의 객체에 입력
                    // * TopName 객체 데이터 입력
                    topName = new TopName();
                    topName.version = sheet2ValueArray[2, 2] == null ? null : sheet2ValueArray[2, 2].ToString();
                    topName.creator = sheet2ValueArray[2, 3] == null ? null : sheet2ValueArray[2, 3].ToString();
                    topName.progress = Convert.ToInt32(sheet2ValueArray[2, 4]);
                    topName.formatt = sheet2ValueArray[2, 5] == null ? null : sheet2ValueArray[2, 5].ToString();
                    topName.time = Convert.ToDouble(sheet2ValueArray[2, 6]);
                    topName.data = new List<object>();

                    // * TopName 객체 내의 Data 객체 리스트 입력
                    IList<object> titleList = new List<object>();
                    for (int r = 2; r <= sheet2ValueArray.GetLength(0); r++)
                    {
                        object tempTitle = sheet2ValueArray[r, 7];
                        if (!titleList.Any())   // 리스트에 아무것도 없을때 (=맨처음)
                        {
                            titleList.Add(tempTitle);
                        }
                        else if (tempTitle == null)  // null 이거나 "" 일 때 tempTitle == ""
                        {
                            titleList.Add(tempTitle);
                        }
                        else if (titleList.Contains(tempTitle)) // 타이틀 이미 입력됨(통과)
                        {
                            continue;
                        }

                        Data tempData = new Data();
                        tempData.title = tempTitle == null ? "" : tempTitle.ToString();
                        tempData.paragraphs = new List<object>();

                        topName.data.Add(tempData);
                    }

                    // * TopName->Data 객체 리스트 내의 Paragraphs 객체 리스트 입력
                    int dataCount = 0;
                    object currentTitle = sheet2ValueArray[2, 7];
                    List<Data> tempDataList = topName.data.Cast<Data>().ToList();
                    for (int r = 2; r <= sheet2ValueArray.GetLength(0); r++)
                    {
                        Paragraphs tempParagraphs = new Paragraphs();
                        tempParagraphs.context = sheet2ValueArray[r, 8] == null ? null : sheet2ValueArray[r, 8].ToString();
                        //tempParagraphs.context_original = sheet2ValueArray[r, 9] == null ? null : sheet2ValueArray[r, 9].ToString();
                        tempParagraphs.context_en = sheet2ValueArray[r, 9] == null ? null : sheet2ValueArray[r, 9].ToString();
                        tempParagraphs.context_tagged = sheet2ValueArray[r, 10] == null ? null : sheet2ValueArray[r, 10].ToString();
                        //if (sheet2ValueArray[r, 11] == null)
                        //{
                        //    tempParagraphs.context_tagged = null;
                        //}
                        //else
                        //{
                        //    //tempParagraphs.context_tagged = new List<string>();
                        //    string[] tempTagged = sheet2ValueArray[r, 11].ToString().Split(':');
                        //    foreach (var item in tempTagged)
                        //    {
                        //        tempParagraphs.context_tagged.Add(item);
                        //    }
                        //}
                        tempParagraphs.qas = new List<object>();

                        if (sheet2ValueArray[r, 7] == null || sheet2ValueArray[r, 7].ToString() == "")
                        {
                            if (r != 2)
                            {
                                dataCount++;
                            }
                            tempDataList[dataCount].paragraphs.Add(tempParagraphs);
                            currentTitle = sheet2ValueArray[r, 7] == null ? null : sheet2ValueArray[r, 7].ToString();
                        }
                        else if (sheet2ValueArray[r, 7] == currentTitle)
                        {
                            tempDataList[dataCount].paragraphs.Add(tempParagraphs);
                        }
                        else
                        {
                            dataCount++;
                            tempDataList[dataCount].paragraphs.Add(tempParagraphs);
                            currentTitle = sheet2ValueArray[r, 7].ToString();
                        }
                    }
                    topName.data = tempDataList.Cast<object>().ToList();

                    // * TopName->Data->Paragraphs 객체 리스트 내의 Qas 객체 리스트 입력
                    dataCount = 0;
                    int paragraphCount = 0;
                    int currentParagraph = 1;
                    tempDataList = topName.data.Cast<Data>().ToList();
                    List<Qas> tempQasList = new List<Qas>();
                    for (int r = 2; r <= sheet1ValueArray.GetLength(0); r++)
                    {
                        Qas tempQas = new Qas();
                        tempQas.id = sheet1ValueArray[r, 2] == null ? null : sheet1ValueArray[r, 2].ToString();
                        tempQas.confuseQt = Convert.ToBoolean(sheet1ValueArray[r, 3] == null ? null : sheet1ValueArray[r, 3]);
                        tempQas.confuseQf = Convert.ToBoolean(sheet1ValueArray[r, 4] == null ? null : sheet1ValueArray[r, 4]);
                        tempQas.confuseLat = Convert.ToBoolean(sheet1ValueArray[r, 5] == null ? null : sheet1ValueArray[r, 5]);
                        tempQas.confuseSat = Convert.ToBoolean(sheet1ValueArray[r, 6] == null ? null : sheet1ValueArray[r, 6]);
                        
                        tempQas.question = sheet1ValueArray[r, 7] == null ? null : sheet1ValueArray[r, 7].ToString();
                       // tempQas.question_original = sheet1ValueArray[r, 5] == null ? null : sheet1ValueArray[r, 5].ToString();
                        tempQas.question_en = sheet1ValueArray[r, 8] == null ? null : sheet1ValueArray[r, 8].ToString();
                        tempQas.question_tagged = sheet1ValueArray[r, 9] == null ? null : sheet1ValueArray[r, 9].ToString();
                        
                        //if (sheet1ValueArray[r, 7] == null)
                        //{
                        //    tempQas.question_tagged = null;
                        //}
                        //else
                        //{
                        //    tempQas.question_tagged = new List<string>();
                        //    string[] tempTagged = sheet1ValueArray[r, 7].ToString().Split(':');
                        //    foreach (var item in tempTagged)
                        //    {
                        //        tempQas.question_tagged.Add(item);
                        //    }
                        //}
                        tempQas.questionType = sheet1ValueArray[r, 10] == null ? null : sheet1ValueArray[r, 10].ToString();
                        tempQas.questionFocus = sheet1ValueArray[r, 11] == null ? null : sheet1ValueArray[r, 11].ToString();
                        tempQas.questionSAT = sheet1ValueArray[r, 12] == null ? null : sheet1ValueArray[r, 12].ToString();
                        tempQas.questionLAT = sheet1ValueArray[r, 13] == null ? null : sheet1ValueArray[r, 13].ToString();

                        tempQas.etriQtCheck = Convert.ToBoolean(sheet1ValueArray[r, 14] == null ? null : sheet1ValueArray[r, 14]);
                        tempQas.etriQfCheck = Convert.ToBoolean(sheet1ValueArray[r, 15] == null ? null : sheet1ValueArray[r, 15]);
                        tempQas.etriLatCheck = Convert.ToBoolean(sheet1ValueArray[r, 16] == null ? null : sheet1ValueArray[r, 16]);
                        tempQas.etriSatCheck = Convert.ToBoolean(sheet1ValueArray[r, 17] == null ? null : sheet1ValueArray[r, 17]);

                        tempQas.etriQt = sheet1ValueArray[r, 18] == null ? null : sheet1ValueArray[r, 18].ToString();//
                        tempQas.etriQf = sheet1ValueArray[r, 19] == null ? null : sheet1ValueArray[r, 19].ToString();//
                        tempQas.etriLat = sheet1ValueArray[r, 20] == null ? null : sheet1ValueArray[r, 20].ToString();//
                        tempQas.etriSat = sheet1ValueArray[r, 21] == null ? null : sheet1ValueArray[r, 21].ToString();//
                        List<Answers> tempAnswersList = new List<Answers>();

                        // * TopName->Data->Paragraphs->Qas 객체 리스트 내의 Answers 객체 리스트 입력
                        for (int i = 0; i < 3; i++)
                        {
                            int ansStartColNum = 22 + (i * 6);//18
                            if (sheet1ValueArray[r, ansStartColNum] == null)
                            {
                                break;      // 정답의 text 공백이면 없음 처리
                            }

                            Answers tempAnswers = new Answers();
                            tempAnswers.text = sheet1ValueArray[r, ansStartColNum] == null ? null : sheet1ValueArray[r, ansStartColNum].ToString();
                            //tempAnswers.text_original = sheet1ValueArray[r, ansStartColNum + 1] == null ? null : sheet1ValueArray[r, ansStartColNum + 1].ToString();
                            tempAnswers.text_en = sheet1ValueArray[r, ansStartColNum + 1] == null ? null : sheet1ValueArray[r, ansStartColNum + 1].ToString();
                            tempAnswers.text_tagged = sheet1ValueArray[r, ansStartColNum + 2] == null ? null : sheet1ValueArray[r, ansStartColNum + 2].ToString();
                            tempAnswers.text_syn = sheet1ValueArray[r, ansStartColNum + 3] == null ? null : sheet1ValueArray[r, ansStartColNum + 3].ToString();
                            //if (sheet1ValueArray[r, ansStartColNum + 3] == null)
                            //{
                            //    tempAnswers.text_tagged = null;
                            //}
                            //else
                            //{
                            //    tempAnswers.text_tagged = new List<string>();
                            //    string[] tempTagged = sheet1ValueArray[r, ansStartColNum + 3].ToString().Split(':');
                            //    foreach (var item in tempTagged)
                            //    {
                            //        tempAnswers.text_tagged.Add(item);
                            //    }
                            //}
                            //if (sheet1ValueArray[r, ansStartColNum + 4] == null)
                            //{
                            //    tempAnswers.text_syn = null;
                            //}
                            //else
                            //{
                            //    tempAnswers.text_syn = new List<string>();
                            //    string[] tempSyn = sheet1ValueArray[r, ansStartColNum + 4].ToString().Split(':');
                            //    foreach (var item in tempSyn)
                            //    {
                            //        tempAnswers.text_syn.Add(item);
                            //    }
                            //}
                            tempAnswers.answer_start = Convert.ToInt32(sheet1ValueArray[r, ansStartColNum + 4]);
                            tempAnswers.answer_end = Convert.ToInt32(sheet1ValueArray[r, ansStartColNum + 5]);

                            tempAnswersList.Add(tempAnswers);
                        }
                        tempQas.answers = tempAnswersList.Cast<object>().ToList();

                        tempQasList.Add(tempQas);
                        currentParagraph = Convert.ToInt32(sheet1ValueArray[r, 40]);//36

                        if (r + 1 <= sheet1ValueArray.GetLength(0)) // 다음 목표 row가 sheet1ValueArray의 1차 배열 길이를 넘지 않을때
                        {
                            if (currentParagraph != Convert.ToInt32(sheet1ValueArray[r + 1, 40]))   // 현재 row의 소속 paragraph 값과 다음 row의 소속 paragraph값을 비교하여 같지 않다면
                            {
                                topName.data.Cast<Data>().ToList()[dataCount].paragraphs.Cast<Paragraphs>().ToList()[paragraphCount].qas = tempQasList.Cast<object>().ToList(); // Qas 리스트 삽입
                                tempQasList = new List<Qas>();
                                if (paragraphCount < topName.data.Cast<Data>().ToList()[dataCount].paragraphs.Count - 1) // paragraphCount 값이 현재 Data에서의 끝에 도달하기 전에는 이렇게 처리
                                {
                                    paragraphCount++;
                                }
                                else    // 도달하고 난 후에는 이렇게 처리
                                {
                                    dataCount++;
                                    paragraphCount = 0;
                                }
                            }
                        }

                        if (r == sheet1ValueArray.GetLength(0))  // 현재 row가 마지막일때
                        {
                            topName.data.Cast<Data>().ToList()[dataCount].paragraphs.Cast<Paragraphs>().ToList()[paragraphCount].qas = tempQasList.Cast<object>().ToList();
                        }
                        
                    }

                    // ** JSON 파일로 저장
                    m_savePath = Path.ChangeExtension(m_path, "json");
                    FileInfo fi = new FileInfo(m_savePath);
                    if (fi.Exists)  // 파일이 이미 존재하면 삭제
                    {
                        fi.Delete();
                    }

                    string saveJSONText;
                    if (m_EtoJNullRemoveCheck)
                    {

                        saveJSONText = JsonConvert.SerializeObject(topName, Formatting.Indented, new JsonSerializerSettings
                            {
                                NullValueHandling = NullValueHandling.Ignore    // Null값 객체 제거
                            }
                            );
                    }
                    else
                    {
                        saveJSONText = JsonConvert.SerializeObject(topName, Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Include   // Null값 객체 포함
                        }
                            );
                    }

                    using (StreamWriter sw = new StreamWriter(m_savePath))
                    {
                        sw.Write(saveJSONText);
                    }

                    #endregion
                }
                return STR_CONVERTING_SUCCESS;
            }
            catch (Exception e)
            {
                if (excelOpen)
                {
                    Marshal.FinalReleaseComObject(range);
                    Marshal.FinalReleaseComObject(objWorksheet);

                    Marshal.FinalReleaseComObject(objWorksheets);

                    objWorkbook.Close(false, missing, missing);
                    objWorkbooks.Close();
                    objApp.Quit();

                    Marshal.FinalReleaseComObject(objWorkbook);
                    Marshal.FinalReleaseComObject(objWorkbooks);
                    Marshal.FinalReleaseComObject(objApp);

                    objApp = null;
                }

                return "예외처리 된 오류 발생.\r\n파일: " + filePath;
            }
        }
    }
}
