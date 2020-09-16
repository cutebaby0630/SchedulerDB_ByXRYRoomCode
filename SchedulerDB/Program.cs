using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SqlServerHelper.Core;
using SqlServerHelper;
using System.ComponentModel;
using LicenseContext = OfficeOpenXml.LicenseContext;
using System.Reflection;
using System.ComponentModel.DataAnnotations;

namespace SchedulerDB
{
    public static class ExcelExtensions
    {
        // SetQuickStyle，指定前景色/背景色/水平對齊
        public static void SetQuickStyle(this ExcelRange range,
            Color fontColor,
            Color bgColor = default(Color),
            ExcelHorizontalAlignment hAlign = ExcelHorizontalAlignment.Left)
        {
            range.Style.Font.Color.SetColor(fontColor);
            if (bgColor != default(Color))
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid; // 一定要加這行..不然會報錯
                range.Style.Fill.BackgroundColor.SetColor(bgColor);
            }
            range.Style.HorizontalAlignment = hAlign;
        }

        //讓文字上有連結
        public static void SetHyperlink(this ExcelRange range, Uri uri)
        {
            range.Hyperlink = uri;
            range.Style.Font.UnderLine = true;
            range.Style.Font.Color.SetColor(Color.Blue);
        }
    }
    class Program
    {

        static void Main(string[] args)
        {
            var fliepath = $@"C:\Users\v-vyin\SchedulerDB_ExcelFile\{"Scheduler" + DateTime.Now.ToString("yyyyMMddhhmm")}";
            Directory.CreateDirectory(fliepath);
            string sql = @"
IF OBJECT_ID('TEMPDB..#Source1') is not null DROP TABLE #Source1
SELECT A.*

,ROW_NUMBER() OVER (PARTITION BY DVC_CHRT,DVC_RQNO,DVC_ROOM ORDER BY [Priority] ASC, ModifyTime DESC ) AS ROWNO 
INTO #Source1 
FROM (
SELECT 
       x.DVC_ROOM,x.DVC_CHRT,x.DVC_RQNO,x.DVC_DATE,x.DVC_STTM,
              x.DVC_CON1,x.DVC_CON2,x.DVC_CON3 ,x.DVC_TXDT+x.DVC_TXTM ModifyTime
          ,1 AS [Priority]
       ,'XRYMDVCF' SourceTable
FROM PXRYDB.SKDBA.XRYMDVCF X
WHERE DVC_CHRT <> ''
UNION ALL
SELECT 
          OPE_OPRM,OPE_CHRT,OPE_ODRN,OPE_DATE,OPE_TIME,
          OPE_OPNO,'','',OPE_TXDT+OPE_TXTM ModifyTime
         
         ,2 AS [Priority]
      ,'OPDMOPEF' SourceTable
FROM POPDDB.SKDBA.OPDMOPEF
where OPE_NOTE = '' --非取消

UNION ALL
SELECT 
          CAT_ROOM,CAT_CHRT,CAT_RQNO,CAT_DATE,CAT_TIME,
          '','','',CAT_DATE+CAT_STTM ModifyTime
         ,3 AS [Priority]
      ,'XRYMCATF' SourceTable
FROM PXRYDB.SKDBA.XRYMCATF

UNION ALL
SELECT 
          'DS',REC_CHAT,ODD_ODRN,REC_DAT7,REC_TIME
      ,substring(ODD_PKEY,11,7),'','',REC_DAT7+REC_TIME ModifyTime
          ,4 AS [Priority]
      ,'OPDTODDF' SourceTable
FROM POPDDB.SKDBA.opdtoddf 
 INNER JOIN POPDDB.SKDBA.OPDTRECF ON REC_PKEY=SUBSTRING(ODD_PKEY,1,10)
WHERE odd_stck = 'DS'
AND ODD_ODRN<>''


UNION ALL
SELECT 
         LEFT(MWE_SCDT,2) MWE_ROOM,left(a.MWE_PKEY,8) MWE_CHRT,SUBSTRING(a.MWE_PKEY,9,9) MWE_RQNO, MWE_CHD7,MWE_CHTM,
       a.MWE_OPNO,'','',MWE_CHD7+MWE_CHTM ModifyTime
      
          ,5 AS [Priority]
       ,'XRYMMWER' SourceTable
FROM PXRYDB.SKDBA.XRYMMWER a 
WHERE a.MWE_CHD7 >= '1090501'
AND a.MWE_SCTM In (N'7421', N'7423', N'7424', N'7425', N'7426')
AND a.MWE_EXDT = ''
) A


--SELECT * FROM #Source1 WHERE ROWNO = 1


IF object_id('tempdb..#RESTTReservation') IS NOT NULL DROP TABLE #RESTTReservation

DECLARE @DVC_DATE char(4) = '0916'

SELECT a.ReservationId,a.RoomCode RESRoomCode,
       b.DVC_ROOM XRYRoomCode,a.CalendarGroupName,
       a.MedicalNoteNo,a.ApplyFormNo,a.[Start],a.SourceCode,a.SourceTable,
          b.DVC_CHRT,b.DVC_RQNO,b.DVC_DATE,b.DVC_STTM,b.SourceTable XRYSourceCode,
          b.DVC_CON1,b.DVC_CON2,b.DVC_CON3,b.ModifyTime
INTO #RESTTReservation
FROM 
(
       SELECT distinct
                     --a.id, cal.CalendarCode RoomCode,calg.DisplayName CalendarGroupName,a.MedicalNoteNo,
                 a.ReservationId, g.RoomCode RoomCode,g.CalendarGroupName ,a.MedicalNoteNo,
              a.ApplyFormNo,c.[Start],a.SourceCode,
               'RESTTReservation' SourceTable
       FROM HISSCHDB.dbo.RESTReservationOrder a
       INNER JOIN HISSCHDB.dbo.RESTTimeslotRes b ON a.ReservationId = b.ReservationId
       INNER JOIN HISSCHDB.dbo.RESTTimeslot c ON b.TimeslotId = c.Id
          left JOIN HISSCHDB.dbo.RESTReservationOrderDetail d ON a.ReservationId = d.ReservationId
          left JOIN HISSCHDB.dbo.RESTReservationOrderDetailForEXA e ON d.ReservationDetailId = e.ReservationDetailId
          left JOIN HISDB.dbo.EXATOrderDetail f ON e.ExaOrderDetailId = f.ExaOrderDetailId
          LEFT JOIN HISSCHDB.dbo.tmpEXAMRoomMapping g ON f.OldRoomCode = g.OldRoomCode
       --LEFT JOIN HISSCHDB.dbo.PROMCalendar cal ON a.CalendarId = cal.Id 
       --LEFT JOIN HISSCHDB.dbo.PROMCalendarGroup calg ON cal.CalendarGroupId = calg.Id
       --LEFT JOIN (SELECT *,ROW_NUMBER() OVER (PARTITION BY CalendarId ORDER BY RoomCode) rownum
       --           FROM HISSCHDB.dbo.tmpEXAMRoomMapping 
       --                 WHERE OldRoomCode IS NOT NULL ) d ON a.CalendarId = d.CalendarId AND rownum = 1
       WHERE a.MedicalOrderCode = 'EXA'
       --AND convert(date,c.[Start]) = '2020'+@DVC_DATE
       --ORDER BY d.RoomCode,c.[Start]
) a
FULL OUTER JOIN #Source1 b ON a.MedicalNoteNo = b.DVC_CHRT AND a.ApplyFormNo = b.DVC_RQNO

SELECT RESRoomCode,XRYRoomCode,CalendarGroupName,MedicalNoteNo,Start,DVC_CHRT,DVC_RQNO,DVC_DATE,DVC_STTM 
FROM #RESTTReservation a
ORDER by XRYRoomCode, RESRoomCode
";

            //Step 1.讀取DB Table List
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsetting.json", optional: true, reloadOnChange: true).Build();
            //取得連線字串
            string connString = config.GetConnectionString("DefaultConnection");
            //string connString = "Data Source=10.1.222.181;Initial Catalog={0};Integrated Security=False;User ID={1};Password={2};Pooling=True;MultipleActiveResultSets=True;Connect Timeout=120;Encrypt=False;TrustServerCertificate=True;ApplicationIntent=ReadWrite";
            SqlServerDBHelper sqlHelper = new SqlServerDBHelper(string.Format(connString, "HISDB", "msdba", "1qaz@wsx"));
            // DataTable dt = sqlHelper.FillTableAsync(sql).Result;
            DataTable dt = sqlHelper.FillTableAsync(sql).Result;
            //印出list
            //int rowCount = (dt == null) ? 0 : dt.Rows.Count;
            //Console.WriteLine(rowCount);

            //Step 1.1.將資料放入List
            List<DBData> migrationTableInfoList = sqlHelper.QueryAsync<DBData>(sql).Result?.ToList();
            //Step 1.2 將date Distinct排序給sheet用 > 遞增 order by 遞減OrderByDescending
            var datetime = migrationTableInfoList.Where(p => p.Start != DateTime.MinValue ? p.Start.Date >= DateTime.Now.AddDays(-2) && p.Start.Date <= DateTime.Now.AddDays(14) : p.PlanDate.Date >= DateTime.Now.AddDays(-1) && p.PlanDate.Date <= DateTime.Now.AddDays(14))
                                                 .Select(p => p.Start != DateTime.MinValue ? p.Start.Date : p.PlanDate.Date)
                                                 .OrderBy(p => p.Date)
                                                 .Distinct()
                                                 .ToList();
            //以科室名稱作為檔案名稱
            var XRYRoomCode = migrationTableInfoList.OrderBy(p => p.XRYRoomCode)
                                                    .Select(p => p.XRYRoomCode == null || p.XRYRoomCode == "" || p.XRYRoomCode == " " || p.XRYRoomCode == "    " ||p.XRYRoomCode =="  " ? "Blank" : p.XRYRoomCode)
                                                    .Distinct()
                                                    .ToList();

            //Step 2.建立 各日期Sheet
            // var excelname = "Scheduler" + DateTime.Now.ToString("yyyyMMddhhmm") + ".xlsx";
            foreach (var date in datetime)
            {
                var excelname = new FileInfo(date.ToString("yyyyMMdd") + ".xlsx");
                //ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excel = new ExcelPackage(excelname))
                {
                    var importDBData = new ImportDBData();
                    importDBData.GenFirstSheet(excel, XRYRoomCode);
                    for (int sheetnum = 0; sheetnum <= XRYRoomCode.Count - 1; sheetnum++)
                    {
                        //Step 3.將對應的List 丟到各Sheet中
                        ExcelWorksheet sheet = excel.Workbook.Worksheets.Add(XRYRoomCode[sheetnum]);
                        //抽function
                        int rowIndex = 2;
                        int colIndex = 1;
                        importDBData.ImportData(dt, sheet, rowIndex, colIndex, migrationTableInfoList, date);
                    }
                    // Step 4.Export EXCEL
                    Byte[] bin = excel.GetAsByteArray();
                    File.WriteAllBytes(fliepath.ToString() + @"\" + excelname, bin);

                }
            }


            //Step 5. Send Email
            var helper = new SMTPHelper("lovemath0630@gmail.com", "koormyktfbbacpmj", "smtp.gmail.com", 587, true, true); //寄出信email
            string subject = $"Datebase Scheduler報表 {DateTime.Now.ToString("yyyyMMdd")}"; //信件主旨
            string body = $"Hi All, \r\n\r\n{DateTime.Now.ToString("yyyyMMdd")} Scheduler報表 如附件，\r\n\r\n Best Regards, \r\n\r\n Vicky Yin";//信件內容
            string attachments = null;//附件
            var fileName = @"C:\Users\v-vyin\SchedulerDB_ExcelFile\";//附件位置
            if (File.Exists(fileName.ToString()))
            {
                attachments = fileName.ToString();
            }
            string toMailList = "lovemath0630@gmail.com;v-vyin@microsoft.com";//收件者
            string ccMailList = "";//CC收件者

            helper.SendMail(toMailList, ccMailList, null, subject, body, null);
        }

        #region --DBdata getset
        public class DBData
        {
            [Required]
            [DisplayName("新檢查室")]
            public string RESRoomCode { get; set; }
            [Required]
            [DisplayName("舊檢查室")]
            public string XRYRoomCode { get; set; }
            [Required]
            [DisplayName("檢查室名稱")]
            public string CalendarGroupName { get; set; }
            [Required]
            [DisplayName("病歷號")]
            public string MedicalNoteNo { get; set; }
            [Required]
            [DisplayName("檢查單號")]
            public string ExaRequestNo { get; set; }
            [Required]
            [DisplayName("檢查時間")]
            public DateTime Start { get; set; }
            public string ReservationSourceType { get; set; }
            public string SourceCode { get; set; }
            [Required]
            [DisplayName("主機病歷號")]
            public string DVC_CHRT { get; set; }
            [Required]
            [DisplayName("主機單號")]
            public string DVC_RQNO { get; set; }
            [Required]
            [DisplayName("主機排程日")]
            public string DVC_DATE { get; set; }
            [Required]
            [DisplayName("主機排程時間")]
            public string DVC_STTM { get; set; }
            [Required]
            [DisplayName("主機檢查碼1")]
            public string XRYSourceCode { get; set; }
            public DateTime PlanDate { get; set; }
        }
        #endregion
        #region -- Data to excel
        public class ImportDBData
        {
            private ExcelWorksheet _sheet { get; set; }
            private int _rowIndex { get; set; }
            private int _colIndex { get; set; }
            private DataTable _dt { get; set; }
            private List<DBData> _dblist { get; set; }
            public void ImportData(DataTable dt, ExcelWorksheet sheet, int rowIndex, int colIndex, List<DBData> dblist, DateTime date)
            {
                _sheet = sheet;
                _rowIndex = rowIndex;
                _colIndex = colIndex;
                _dt = dt;
                _dblist = dblist;
                _sheet.Cells[_rowIndex - 1, _colIndex].Value = "返回目錄";
                _sheet.Cells[_rowIndex - 1, _colIndex].SetHyperlink(new Uri($"#'目錄'!A1", UriKind.Relative));
                string temp_MedicalNoteNo = null;
                string temp_ExaRequestNo = null;
                string temp_DVC_CHRT = null;
                //3.1塞columnName 到Row 
                for (int columnNameIndex = 0; columnNameIndex <= _dt.Columns.Count - 1; columnNameIndex++)
                {
                    MemberInfo property = typeof(DBData).GetProperty((_dt.Columns[columnNameIndex].ColumnName == null ? string.Empty : _dt.Columns[columnNameIndex].ColumnName));
                    var attribute = property.GetCustomAttributes(typeof(DisplayNameAttribute), true)
                                            .Cast<DisplayNameAttribute>().Single();
                    string columnName = attribute.DisplayName;
                    _sheet.Cells[_rowIndex, _colIndex++].Value = columnName;


                }
                _sheet.Cells[_rowIndex, 1, _rowIndex, _colIndex - 1]
                     .SetQuickStyle(Color.Black, Color.LightPink, ExcelHorizontalAlignment.Center);

                //將對應值放入
                foreach (var dbdata in _dblist)
                {
                    if (_sheet.ToString() == (dbdata.XRYRoomCode == null ? "Blank" : dbdata.XRYRoomCode) && date.ToString("yyyy-MM-dd") == (dbdata.Start != DateTime.MinValue ? dbdata.Start.ToString("yyyy-MM-dd") : dbdata.PlanDate.ToString("yyyy-MM-dd")))
                    {
                        _rowIndex++;
                        _colIndex = 1;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.RESRoomCode;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.XRYRoomCode;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.CalendarGroupName;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.MedicalNoteNo;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.ExaRequestNo;
                        _sheet.Cells[_rowIndex, _colIndex].Value = dbdata.Start;
                        _sheet.Cells[_rowIndex, _colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        _sheet.Cells[_rowIndex, _colIndex++].Style.Numberformat.Format = "yyyy/MM/dd HH:mm:ss";
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_CHRT;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_RQNO;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_DATE;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_STTM;
                        if (dbdata.MedicalNoteNo == (temp_MedicalNoteNo == null ? string.Empty : temp_MedicalNoteNo) && dbdata.ExaRequestNo == (temp_ExaRequestNo == null ? string.Empty : temp_ExaRequestNo) && dbdata.DVC_CHRT == (temp_DVC_CHRT == null ? string.Empty : temp_DVC_CHRT))
                        {
                            _sheet.Cells[_rowIndex--, _colIndex].Value = "v";
                            _sheet.Cells[_rowIndex++, _colIndex].Value = "v";
                        }
                        temp_MedicalNoteNo = dbdata.MedicalNoteNo;
                        temp_ExaRequestNo = dbdata.ExaRequestNo;
                        temp_DVC_CHRT = dbdata.DVC_CHRT;
                    }
                }

                //Autofit
                int startColumn = _sheet.Dimension.Start.Column;
                int endColumn = _sheet.Dimension.End.Column;
                for (int count = startColumn; count <= endColumn; count++)
                {
                    _sheet.Column(count).AutoFit();
                }


            }
            public void GenFirstSheet(ExcelPackage excel, List<string> list)
            {
                int rowIndex = 1;
                int colIndex = 1;

                int maxCol = 0;

                ExcelWorksheet firstSheet = excel.Workbook.Worksheets.Add("目錄");

                firstSheet.Cells[rowIndex, colIndex++].Value = "";
                firstSheet.Cells[rowIndex, colIndex++].Value = "檢查時間";

                firstSheet.Cells[rowIndex, 1, rowIndex, colIndex - 1]
                    .SetQuickStyle(Color.Black, Color.LightPink, ExcelHorizontalAlignment.Center);

                maxCol = Math.Max(maxCol, colIndex - 1);

                foreach (string info in list)
                {
                    rowIndex++;
                    colIndex = 1;

                    firstSheet.Cells[rowIndex, colIndex++].Value = rowIndex - 1;
                    firstSheet.Cells[rowIndex, colIndex++].Value = info;
                    firstSheet.Cells[rowIndex, colIndex - 1].SetHyperlink(new Uri($"#'{(string.IsNullOrEmpty(info) ? "Blank" : info)}'!A1", UriKind.Relative));
                }

                for (int i = 1; i <= maxCol; i++)
                {
                    firstSheet.Column(i).AutoFit();
                }
            }

        }
        #endregion

    }
}
