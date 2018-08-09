using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Runtime.InteropServices;


using System.Drawing;
using JSON_ExcelDirectionalConverter.CtagClasses;

namespace JSON_ExcelDirectionalConverter
{
    class CrossJEConverter
    {
        const string STR_CONVERTING_SUCCESS = "SUCCESS";

        private convertingMode m_currentConvertingMode;
        private string m_path;
        private string m_savePath;
        private string m_saveFileName;
        private int m_fileCount;
        bool m_wholeSaveinOneFile;

        //private bool m_EtoJNullRemoveCheck;

        private static string[] sheet1ColHeader = {"count", "id", "confuseQt1", "confuseQf1", "confuseSat1", "confuseLat1", "question", "question_en", "question_tagged1", "questionType1",
                                                  "questionFocus1", "questionSAT1", "questionLAT1", "confuseQt2", "confuseQf2", "confuseSat2", "confuseLat2", "question_tagged2",
                                                  "questionType2", "questionFocus2", "questionSAT2", "questionLAT2", "confuseQt3", "confuseQf3", "confuseSat3", "confuseLat3", "question_tagged3",
                                                  "questionType3", "questionFocus3", "questionSAT3", "questionLAT3", "text", "text_en", "text_tagged", "text_syn", "answer_start",
                                                  "answer_end", "para_group_num" ,"file_group_num"};

        private int sheet1ColCount = sheet1ColHeader.Length;
        private int sheet1RowCount;

        private static string[] sheet2ColHeader = {"count", "version", "creator", "progress", "formatt", "time", "check", "firstfile", "secondfile",
                                                   "title",
                                                  "context", "context_en", "context_tagged"};

        private int sheet2ColCount = sheet2ColHeader.Length;
        private int sheet2RowCount;
        
        private Excel.Application objApp;
        private Excel.Workbooks objWorkbooks;
        private Excel.Workbook objWorkbook;
        private Excel.Sheets objWorksheets;
        private Excel.Worksheet objWorksheet;
        private Excel.Range range;

        private Excel.XlHAlign HCENTER = Excel.XlHAlign.xlHAlignCenter;

        public CrossJEConverter(convertingMode mode, string saveFileName, bool wholeSaveinOneFiles)//, bool EtoJNullRemoveCheck)
        {
            m_currentConvertingMode = mode;
            m_saveFileName = saveFileName;
            m_wholeSaveinOneFile = wholeSaveinOneFiles;
            sheet1RowCount = 0;
            sheet2RowCount = 0;
            //m_EtoJNullRemoveCheck = EtoJNullRemoveCheck;
        }

        public string convertFiles(IList<string> filePaths)
        {
            m_fileCount = filePaths.Count;
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

            Cross_TopTag topTag;
            Cross_Data[] data;
            Cross_Paragraphs[][] paragraphs;
            Cross_Qas[][][] qas;
            Cross_Answers[][][][] answers;

            object[,] sheet1ValueArray;
            object[,] sheet2ValueArray;

            int countParagraphs = 0;
            int countQas = 0;
            int currentRow = 0;

            bool excelOpen = false;

            try
            {

                if (m_currentConvertingMode == convertingMode.CJSONToCExcel)
                {
                    #region JSON -> Excel 변환

                    // ** name1 영역 파싱
                    topTag = JsonConvert.DeserializeObject<Cross_TopTag>(File.ReadAllText(m_path));

                    // name2 영역 파싱
                    data = new Cross_Data[topTag.data.Count];
                    for (int i = 0; i < data.Length; i++)
                    {
                        data[i] = JsonConvert.DeserializeObject<Cross_Data>(topTag.data[i].ToString());
                    }

                    // ** name3 영역 파싱
                    paragraphs = new Cross_Paragraphs[data.Length][];
                    for (int i = 0; i < data.Length; i++)
                    {
                        paragraphs[i] = new Cross_Paragraphs[data[i].paragraphs.Count];
                        for (int j = 0; j < data[i].paragraphs.Count; j++)
                        {
                            paragraphs[i][j] = JsonConvert.DeserializeObject<Cross_Paragraphs>(data[i].paragraphs[j].ToString());
                            countParagraphs++;
                        }
                    }

                    // ** name4 영역 파싱
                    qas = new Cross_Qas[data.Length][][];
                    for (int i = 0; i < data.Length; i++)
                    {
                        qas[i] = new Cross_Qas[paragraphs[i].Length][];
                        for (int j = 0; j < paragraphs[i].Length; j++)
                        {
                            qas[i][j] = new Cross_Qas[paragraphs[i][j].qas.Count];
                            for (int k = 0; k < paragraphs[i][j].qas.Count; k++)
                            {
                                qas[i][j][k] = JsonConvert.DeserializeObject<Cross_Qas>(paragraphs[i][j].qas[k].ToString());
                                countQas++;
                            }
                        }
                    }

                    // ** name5 영역 파싱
                    answers = new Cross_Answers[data.Length][][][];
                    for (int i = 0; i < data.Length; i++)
                    {
                        answers[i] = new Cross_Answers[paragraphs[i].Length][][];
                        for (int j = 0; j < paragraphs[i].Length; j++)
                        {
                            answers[i][j] = new Cross_Answers[qas[i][j].Length][];
                            for (int k = 0; k < qas[i][j].Length; k++)
                            {
                                answers[i][j][k] = new Cross_Answers[qas[i][j][k].answers.Count];
                                for (int m = 0; m < qas[i][j][k].answers.Count; m++)
                                {
                                    answers[i][j][k][m] = JsonConvert.DeserializeObject<Cross_Answers>(qas[i][j][k].answers[m].ToString());
                                }
                            }
                        }
                    }

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
                        sheet2ValueArray[row, 1] = topTag.version;
                        sheet2ValueArray[row, 2] = topTag.creator;
                        sheet2ValueArray[row, 3] = topTag.progress;
                        sheet2ValueArray[row, 4] = topTag.formatt;
                        sheet2ValueArray[row, 5] = topTag.time;
                        sheet2ValueArray[row, 6] = topTag.check;
                        sheet2ValueArray[row, 7] = topTag.firstfile;
                        sheet2ValueArray[row, 8] = topTag.secondfile;
                    }

                    // * name2 & name3 영역
                    currentRow = 0;
                    for (int d = 0; d < data.Length; d++)
                    {
                        for (int p = 0; p < paragraphs[d].Length; p++)
                        {
                            sheet2ValueArray[currentRow, 9] = data[d].title;
                            sheet2ValueArray[currentRow, 10] = paragraphs[d][p].context;
                            sheet2ValueArray[currentRow, 11] = paragraphs[d][p].context_en;
                            sheet2ValueArray[currentRow, 12] = paragraphs[d][p].context_tagged;

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
                                sheet1ValueArray[currentRow, 2] = qas[d][p][q].confuseQt1;
                                sheet1ValueArray[currentRow, 3] = qas[d][p][q].confuseQf1;
                                sheet1ValueArray[currentRow, 4] = qas[d][p][q].confuseSat1;
                                sheet1ValueArray[currentRow, 5] = qas[d][p][q].confuseLat1;
                                sheet1ValueArray[currentRow, 6] = qas[d][p][q].question;
                                sheet1ValueArray[currentRow, 7] = qas[d][p][q].question_en;
                                sheet1ValueArray[currentRow, 8] = qas[d][p][q].question_tagged1;
                                sheet1ValueArray[currentRow, 9] = qas[d][p][q].questionType1;
                                sheet1ValueArray[currentRow, 10] = qas[d][p][q].questionFocus1;
                                sheet1ValueArray[currentRow, 11] = qas[d][p][q].questionSAT1;
                                sheet1ValueArray[currentRow, 12] = qas[d][p][q].questionLAT1;
                                sheet1ValueArray[currentRow, 13] = qas[d][p][q].confuseQt2;
                                sheet1ValueArray[currentRow, 14] = qas[d][p][q].confuseQf2;
                                sheet1ValueArray[currentRow, 15] = qas[d][p][q].confuseSat2;
                                sheet1ValueArray[currentRow, 16] = qas[d][p][q].confuseLat2;
                                sheet1ValueArray[currentRow, 17] = qas[d][p][q].question_tagged2;//
                                sheet1ValueArray[currentRow, 18] = qas[d][p][q].questionType2;//
                                sheet1ValueArray[currentRow, 19] = qas[d][p][q].questionFocus2;//
                                sheet1ValueArray[currentRow, 20] = qas[d][p][q].questionSAT2;//
                                sheet1ValueArray[currentRow, 21] = qas[d][p][q].questionLAT2;
                                sheet1ValueArray[currentRow, 22] = qas[d][p][q].confuseQt3;
                                sheet1ValueArray[currentRow, 23] = qas[d][p][q].confuseQf3;
                                sheet1ValueArray[currentRow, 24] = qas[d][p][q].confuseSat3;
                                sheet1ValueArray[currentRow, 25] = qas[d][p][q].confuseLat3;
                                sheet1ValueArray[currentRow, 26] = qas[d][p][q].question_tagged3;
                                sheet1ValueArray[currentRow, 27] = qas[d][p][q].questionType3;
                                sheet1ValueArray[currentRow, 28] = qas[d][p][q].questionFocus3;
                                sheet1ValueArray[currentRow, 29] = qas[d][p][q].questionSAT3;
                                sheet1ValueArray[currentRow, 30] = qas[d][p][q].questionLAT3;

                                sheet1ValueArray[currentRow, 37] = currentParaNum;
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

                                int answerStartColNum = 31;
                                for (int a = 0; a < answers[d][p][q].Length; a++)
                                {
                                    sheet1ValueArray[currentRow, answerStartColNum] = answers[d][p][q][a].text;
                                    sheet1ValueArray[currentRow, answerStartColNum + 1] = answers[d][p][q][a].text_en;
                                    sheet1ValueArray[currentRow, answerStartColNum + 2] = answers[d][p][q][a].text_tagged;
                                    sheet1ValueArray[currentRow, answerStartColNum + 3] = answers[d][p][q][a].text_syn;
                                    sheet1ValueArray[currentRow, answerStartColNum + 4] = answers[d][p][q][a].answer_start;
                                    sheet1ValueArray[currentRow, answerStartColNum + 5] = answers[d][p][q][a].answer_end;
                                    
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

                    range = objWorksheet.get_Range("A1", "M1");
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

                    range = objWorksheet.get_Range("A1", "AL1");
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
                    // * topTag 객체 데이터 입력
                    topTag = new Cross_TopTag();
                    topTag.version = sheet2ValueArray[2, 2] == null ? null : sheet2ValueArray[2, 2].ToString();
                    topTag.creator = sheet2ValueArray[2, 3] == null ? null : sheet2ValueArray[2, 3].ToString();
                    topTag.progress = Convert.ToInt32(sheet2ValueArray[2, 4]);
                    topTag.formatt = sheet2ValueArray[2, 5] == null ? null : sheet2ValueArray[2, 5].ToString();
                    topTag.time = Convert.ToDouble(sheet2ValueArray[2, 6]);
                    topTag.data = new List<object>();

                    // * topTag 객체 내의 Data 객체 리스트 입력
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

                        Cross_Data tempData = new Cross_Data();
                        tempData.title = tempTitle == null ? "" : tempTitle.ToString();
                        tempData.paragraphs = new List<object>();

                        topTag.data.Add(tempData);
                    }

                    // * topTag->Data 객체 리스트 내의 Paragraphs 객체 리스트 입력
                    int dataCount = 0;
                    object currentTitle = sheet2ValueArray[2, 7];
                    List<Cross_Data> tempDataList = topTag.data.Cast<Cross_Data>().ToList();
                    for (int r = 2; r <= sheet2ValueArray.GetLength(0); r++)
                    {
                        Cross_Paragraphs tempParagraphs = new Cross_Paragraphs();
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
                    topTag.data = tempDataList.Cast<object>().ToList();

                    // * topTag->Data->Paragraphs 객체 리스트 내의 Qas 객체 리스트 입력
                    dataCount = 0;
                    int paragraphCount = 0;
                    int currentParagraph = 1;
                    tempDataList = topTag.data.Cast<Cross_Data>().ToList();
                    List<Cross_Qas> tempQasList = new List<Cross_Qas>();
                    for (int r = 2; r <= sheet1ValueArray.GetLength(0); r++)
                    {
                        Cross_Qas tempQas = new Cross_Qas();
                        tempQas.id = sheet1ValueArray[r, 2] == null ? null : sheet1ValueArray[r, 2].ToString();
                        tempQas.confuseQt1 = Convert.ToBoolean(sheet1ValueArray[r, 3] == null ? null : sheet1ValueArray[r, 3]);
                        tempQas.confuseQf1 = Convert.ToBoolean(sheet1ValueArray[r, 4] == null ? null : sheet1ValueArray[r, 4]);
                        tempQas.confuseSat1 = Convert.ToBoolean(sheet1ValueArray[r, 5] == null ? null : sheet1ValueArray[r, 5]);
                        tempQas.confuseLat1 = Convert.ToBoolean(sheet1ValueArray[r, 6] == null ? null : sheet1ValueArray[r, 6]);

                        tempQas.question = sheet1ValueArray[r, 7] == null ? null : sheet1ValueArray[r, 7].ToString();
                        tempQas.question_en = sheet1ValueArray[r, 8] == null ? null : sheet1ValueArray[r, 8].ToString();
                        tempQas.question_tagged1 = sheet1ValueArray[r, 9] == null ? null : sheet1ValueArray[r, 9].ToString();

                        tempQas.questionType1 = sheet1ValueArray[r, 10] == null ? null : sheet1ValueArray[r, 10].ToString();
                        tempQas.questionFocus1 = sheet1ValueArray[r, 11] == null ? null : sheet1ValueArray[r, 11].ToString();
                        tempQas.questionSAT1 = sheet1ValueArray[r, 12] == null ? null : sheet1ValueArray[r, 12].ToString();
                        tempQas.questionLAT1 = sheet1ValueArray[r, 13] == null ? null : sheet1ValueArray[r, 13].ToString();

                        tempQas.confuseQt2 = Convert.ToBoolean(sheet1ValueArray[r, 14] == null ? null : sheet1ValueArray[r, 14]);
                        tempQas.confuseQf2 = Convert.ToBoolean(sheet1ValueArray[r, 15] == null ? null : sheet1ValueArray[r, 15]);
                        tempQas.confuseSat2 = Convert.ToBoolean(sheet1ValueArray[r, 16] == null ? null : sheet1ValueArray[r, 16]);
                        tempQas.confuseLat2 = Convert.ToBoolean(sheet1ValueArray[r, 17] == null ? null : sheet1ValueArray[r, 17]);
                        tempQas.questionType2 = sheet1ValueArray[r, 18] == null ? null : sheet1ValueArray[r, 18].ToString();
                        tempQas.questionFocus2 = sheet1ValueArray[r, 19] == null ? null : sheet1ValueArray[r, 19].ToString();
                        tempQas.questionSAT2 = sheet1ValueArray[r, 20] == null ? null : sheet1ValueArray[r, 20].ToString();
                        tempQas.questionLAT2 = sheet1ValueArray[r, 21] == null ? null : sheet1ValueArray[r, 21].ToString();

                        tempQas.confuseQt3 = Convert.ToBoolean(sheet1ValueArray[r, 22] == null ? null : sheet1ValueArray[r, 22]);
                        tempQas.confuseQf3 = Convert.ToBoolean(sheet1ValueArray[r, 23] == null ? null : sheet1ValueArray[r, 23]);
                        tempQas.confuseSat3 = Convert.ToBoolean(sheet1ValueArray[r, 24] == null ? null : sheet1ValueArray[r, 24]);
                        tempQas.confuseLat3 = Convert.ToBoolean(sheet1ValueArray[r, 25] == null ? null : sheet1ValueArray[r, 25]);
                        tempQas.questionType3 = sheet1ValueArray[r, 26] == null ? null : sheet1ValueArray[r, 26].ToString();
                        tempQas.questionFocus3 = sheet1ValueArray[r, 27] == null ? null : sheet1ValueArray[r, 27].ToString();
                        tempQas.questionSAT3 = sheet1ValueArray[r, 28] == null ? null : sheet1ValueArray[r, 28].ToString();
                        tempQas.questionLAT3 = sheet1ValueArray[r, 29] == null ? null : sheet1ValueArray[r, 29].ToString();

                        List<Cross_Answers> tempAnswersList = new List<Cross_Answers>();

                        // * topTag->Data->Paragraphs->Qas 객체 리스트 내의 Answers 객체 리스트 입력
                        for (int i = 0; i < 3; i++)
                        {
                            int ansStartColNum = 22 + (i * 6);//18
                            if (sheet1ValueArray[r, ansStartColNum] == null)
                            {
                                break;      // 정답의 text 공백이면 없음 처리
                            }

                            Cross_Answers tempAnswers = new Cross_Answers();
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
                                topTag.data.Cast<Cross_Data>().ToList()[dataCount].paragraphs.Cast<Cross_Paragraphs>().ToList()[paragraphCount].qas = tempQasList.Cast<object>().ToList(); // Qas 리스트 삽입
                                tempQasList = new List<Cross_Qas>();
                                if (paragraphCount < topTag.data.Cast<Cross_Data>().ToList()[dataCount].paragraphs.Count - 1) // paragraphCount 값이 현재 Data에서의 끝에 도달하기 전에는 이렇게 처리
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
                            topTag.data.Cast<Cross_Data>().ToList()[dataCount].paragraphs.Cast<Cross_Paragraphs>().ToList()[paragraphCount].qas = tempQasList.Cast<object>().ToList();
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
                    bool m_EtoJNullRemoveCheck = false;
                    if (m_EtoJNullRemoveCheck)
                    {

                        saveJSONText = JsonConvert.SerializeObject(topTag, Formatting.Indented, new JsonSerializerSettings
                        {
                            NullValueHandling = NullValueHandling.Ignore    // Null값 객체 제거
                        }
                            );
                    }
                    else
                    {
                        saveJSONText = JsonConvert.SerializeObject(topTag, Formatting.Indented, new JsonSerializerSettings
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

