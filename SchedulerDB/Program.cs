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
            string sql = @"if object_id('TEMPDB..#XRYTDNLDI1') IS NOT NULL DROP TABLE #XRYTDNLDI1
SELECT DNL_FUNC,DNL_CHRT ,DNL_ODRN,DNL_DATE,DNL_TIME,DNL_ROOM,DNL_ORDR,
       DNL_DEPT,DNL_TXTM,DNL_TEL,DNL_COD1,DNL_RITM,DNL_SOUR
    ,ROW_NUMBER () OVER (PARTITION BY DNL_FUNC, DNL_CHRT,DNL_ODRN ORDER BY DNL_TXTM DESC) AS ROWNO
INTO #XRYTDNLDI1  
FROM PXRYDB.SKDBA.XRYTDNLD 
WHERE DNL_FUNC IN ('I1')
AND DNL_DATE >= '1090501'
--AND DNL_ODRN = '57256661K' --同單號多I1, 取最後一筆為主 by Leon 2020-10-06

--XRYTDNLD取I2與I6
if object_id('TEMPDB..#XRYTDNLD') IS NOT NULL DROP TABLE #XRYTDNLD  
SELECT a.DNL_FUNC,a.DNL_CHRT ,a.DNL_ODRN,a.DNL_DATE,a.DNL_TIME
    ,ROW_NUMBER () OVER (PARTITION BY a.DNL_FUNC, a.DNL_CHRT,a.DNL_ODRN ORDER BY a.DNL_TXTM DESC) AS ROWNO
INTO #XRYTDNLD  
FROM PXRYDB.SKDBA.XRYTDNLD a
INNER HASH JOIN #XRYTDNLDI1 b ON a.DNL_CHRT = b.DNL_CHRT 
                             AND a.DNL_ODRN = b.DNL_ODRN 
                                                                                                 AND a.DNL_TXTM >= b.DNL_TXTM and b.ROWNO = 1 --取最後一筆I1之後的I2&I6 Modify by Leon 2020-10-06
WHERE a.DNL_FUNC IN ('I2','I6')
--and a.DNL_ODRN = '57256661K' 

--XRYTDNLD取D2(取消)
if object_id('TEMPDB..#XRYTDNLD_D2') IS NOT NULL DROP TABLE #XRYTDNLD_D2
SELECT a.DNL_FUNC,a.DNL_CHRT ,a.DNL_ODRN,a.DNL_DATE,a.DNL_TIME
    ,ROW_NUMBER () OVER (PARTITION BY a.DNL_FUNC, a.DNL_CHRT,a.DNL_ODRN ORDER BY a.DNL_TXTM DESC) AS ROWNO
INTO #XRYTDNLD_D2  
FROM PXRYDB.SKDBA.XRYTDNLD a
INNER HASH JOIN #XRYTDNLDI1 b ON a.DNL_CHRT = b.DNL_CHRT 
                             AND a.DNL_ODRN = b.DNL_ODRN 
                                                                                                 AND a.DNL_TXTM >= b.DNL_TXTM and b.ROWNO = 1 --取最後一筆I1之後的I2&I6 Modify by Leon 2020-10-06
WHERE a.DNL_FUNC IN ('D2')

--SELECT * FROM #XRYTDNLDI1 
--SELECT * FROM #XRYTDNLD_D2
--SELECT * FROM #XRYTDNLD
--SELECT * FROM PXRYDB.SKDBA.XRYTDNLD a WHERE a.DNL_ODRN = '57256661K' order by DNL_TXTM





IF object_id('tempdb..#XRYMDVCF1') IS NOT NULL DROP TABLE #XRYMDVCF1

SELECT 
       x.DVC_CHRT,x.DVC_RQNO,rtrim(x.DVC_ROOM) DVC_ROOM,x.DVC_DTNO,
       x.DVC_DEP1,x.DVC_DATE,x.DVC_TXDT,
       x.DVC_TXTM,x.DVC_TXOP,DVC_STTM, 0 AS ModifyEmpId,'' AS DNL_CE_TXDT
                 ,DVC_TEL,
                           CASE WHEN ISNUMERIC(DVC_DATE) = 1 AND ISNUMERIC(DVC_STTM) = 1 
                                         THEN try_convert(datetime,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(DVC_DATE)) = 7 
                                                                     THEN DVC_DATE 
                                                                     ELSE      (CASE WHEN LEFT(DVC_DATE,1) = '0' THEN '1' ELSE '0' END) + DVC_DATE END )) ,1,3)+ 1911)+RIGHT(DVC_DATE,4) +' '+substring(DVC_STTM,1,2)+':'+substring(DVC_STTM,3,2))     
                                         WHEN ISNUMERIC(DVC_DATE) = 1 
                                         THEN try_convert(date,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(DVC_DATE)) = 7 
                                                                     THEN DVC_DATE 
                                                                     ELSE      (CASE WHEN LEFT(DVC_DATE,1) = '0' THEN '1' ELSE '0' END) + DVC_DATE END )),1,3)+ 1911)+RIGHT(DVC_DATE,4))     
                                         END PlanDate
                 ,r.CalendarId
                 ,1 AS [Priority]
       ,'XRYMDVCF' SourceTable
                 ,0 IsDelete
INTO #XRYMDVCF1
FROM PXRYDB.SKDBA.XRYMDVCF x
INNER JOIN hisdb.dbo.PROMMedicalNote n ON x.DVC_CHRT = n.MedicalNoteNo --先過濾病歷號 by Leon
LEFT JOIN hisschdb.dbo.tmpEXAMRoomMapping r ON r.OldRoomCode= DVC_ROOM
WHERE LEN(rtrim(x.DVC_RQNO)) = 9 AND x.DVC_ROOM NOT IN ('575','576')


IF object_id('tempdb..#XRYMPWER') IS NOT NULL DROP TABLE #XRYMPWER

SELECT left(a.PWE_PKEY,8) PWE_CHRT,SUBSTRING(a.PWE_PKEY,9,9) PWE_RQNO, LEFT(PWE_SCDT,2) PWE_ROOM,a.PWE_DTNO,
       a.PWE_DEPT,a.PWE_CHD7,PWE_CHD7 PWE_TXD7,
       PWE_CHTM PWE_TXTM,a.PWE_OPNO,PWE_CHTM, 0 AS ModifyEmpId,'' AS DNL_CE_TXDT
    ,'' PWE_TEL, NULL PlanDate
              ,CASE WHEN PWE_SCNO = N'2071' THEN 166 
                                         WHEN PWE_SCNO = N'3010' AND LEFT(PWE_SCDT,2) In (N'69') THEN 23 
                                         WHEN PWE_SCNO = N'3071' AND LEFT(PWE_SCDT,2) In (N'76') THEN 91 
                                         WHEN PWE_SCNO = N'3072' THEN 91 
                                         WHEN PWE_SCNO = N'3073' THEN 91 
                                         WHEN PWE_SCNO = N'324'  THEN 79
                                         WHEN PWE_SCNO = N'3242' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'02') THEN 21 
                                                                           WHEN LEFT(PWE_SCDT,2) In (N'24') THEN 79 
                                                                                                                                                         WHEN LEFT(PWE_SCDT,2) In (N'83') THEN 80
                                                                                                                                                         ELSE 79 END 
                                         WHEN PWE_SCNO = N'3251' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'02') THEN 21 
                                                                                                                                                         ELSE 79 END 
                                         WHEN PWE_SCNO = N'329' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'29',N'BS') THEN 58 
                                                                           WHEN LEFT(PWE_SCDT,2) In (N'66',N'OC') THEN 59 
                                                                                                                                                         WHEN LEFT(PWE_SCDT,2) In (N'67') THEN 58 
                                                                                                                                                         WHEN LEFT(PWE_SCDT,2) In (N'68') THEN 60
                                                                                                                                                         ELSE 58 END
                                         WHEN PWE_SCNO = N'3321' THEN 77 
                                         WHEN PWE_SCNO = N'3451' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'60') THEN 53 
                                                                           WHEN LEFT(PWE_SCDT,2) In (N'62') THEN 53 
                                                                                                                                                         WHEN LEFT(PWE_SCDT,2) In (N'63') THEN 54
                                                                                                                                                        ELSE 53 END 
                                         WHEN PWE_SCNO = N'3453' THEN 52 
                                         WHEN PWE_SCNO = N'347' THEN 65 
                                         WHEN PWE_SCNO = N'3472' THEN 35 
                                         WHEN PWE_SCNO = N'377' THEN 67 
                                         WHEN PWE_SCNO = N'511' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'08') THEN 61 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'PE') THEN 61 
                                                                                                                                                        ELSE 61 END  
                                         WHEN PWE_SCNO = N'571' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'01',N'06') THEN 25 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'07') THEN 22 
                                                                                                                                                       WHEN LEFT(PWE_SCDT,2) In (N'10') THEN 37 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'12') THEN 35 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'31',N'30') THEN 87 
                                                                                                                                                        ELSE 99 END 
                                         WHEN PWE_SCNO = N'5711' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'08') THEN 61 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'41') THEN 61 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'45') THEN 61 
                                                                                                                                                        ELSE 61 END 
                                         WHEN PWE_SCNO = N'5713' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'30') THEN 89 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'32') THEN 87 
                                                                                                                                                        ELSE 87 END 
                                         WHEN PWE_SCNO = N'5714' THEN 35 
                                         WHEN PWE_SCNO = N'5719' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'41') THEN 88 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'45') THEN 83 
                                                                                                                                                        ELSE 88 END 
                                         WHEN PWE_SCNO = N'572' THEN CASE 
                                                                                                                                                        WHEN substring(PWE_NOTE,2,4) In (N'PAES') THEN 30 --上消化道內視鏡-PAES
                                                                          WHEN substring(PWE_NOTE,2,4) In (N'CES ') THEN 30 --大腸內視鏡-CES
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'SES ') THEN 30 --結腸內視鏡-SES
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'JES ') THEN 30 --小腸內視鏡-JES
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'EES ') THEN 30 --膽道內視鏡超音波-EES
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'ERES') THEN 30 --內視鏡逆行性膽胰攝影-ERES
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'BES ') THEN 34 --支氣管內視鏡-BES
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'ESES') THEN 34 --食道鏡檢查-ESES

                                                                                                                                                       WHEN LEFT(PWE_SCDT,2) In (N'10') THEN CASE WHEN PWE_SRTP = 'P' AND SUBSTRING(PWE_NOTE,2,3) = 'CES' THEN 179
                                                                                                                                                                                                                                                                                                             WHEN PWE_SRTP = 'P' AND SUBSTRING(PWE_NOTE,2,3) = 'PAE' THEN 179
                                                                                                                                                                                                                                                                                                              WHEN PWE_SRTP = 'P' AND SUBSTRING(PWE_NOTE,2,3) = 'SES' THEN 179
                                                                                                                                                                                                                                                                                                             ELSE 30 END
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'11') THEN 34 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'15') THEN 33 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'17') THEN 30 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'18') THEN 30 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'21') THEN 32 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'22') THEN 32 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'23') THEN 32 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'25') THEN 31 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'26') THEN 31 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'27') THEN 31 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'51') THEN 48 
                                                                                                                                                        ELSE 30 END
                                         WHEN PWE_SCNO = N'5721' THEN 48 
                                         WHEN PWE_SCNO = N'5722' THEN 48 
                                         WHEN PWE_SCNO = N'573' THEN CASE WHEN substring(PWE_NOTE,2,4) In (N'AUS ') THEN 93 --腹部超音波-AUS
                                                                          WHEN substring(PWE_NOTE,2,4) In (N'BUS ') THEN 95 --乳房超音波-BUS
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'CUS ') THEN 92 --胸腔超音波-CUS
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'EUS ') THEN 27 --心臟超音波-EUS
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'KUS ') THEN 36 --腎臟超音波-KUS
                                                                                                                                                       WHEN substring(PWE_NOTE,2,4) In (N'TUS ') THEN 37 --甲狀腺超音波TUS 
                                                                                                                                                        WHEN substring(PWE_NOTE,2,4) In (N'THUS') THEN 28 --經直腸攝護腺超音波THUS

                                                                                                                                                       WHEN LEFT(PWE_SCDT,2) In (N'02') THEN 21 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'03') THEN 26 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'05') THEN 27 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'08') THEN 28 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'09') THEN 26 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'10') THEN 93 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'13') THEN 36 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'16') THEN CASE WHEN PWE_SRTP = 'P' AND SUBSTRING(PWE_NOTE,2,3) = 'BUS' THEN 95
                                                                                                                                                                                                                                                                                                             WHEN PWE_SRTP = 'P' AND SUBSTRING(PWE_NOTE,2,3) = 'TUS' THEN 92
                                                                                                                                                                                                                                                                                                             ELSE 92 END
                                                                                                                                                       WHEN LEFT(PWE_SCDT,2) In (N'39') THEN 136 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'51') THEN 48 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'64') THEN 90 
                                                                                                                                                        ELSE 21 END 
                                         WHEN PWE_SCNO = N'5731' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'02') THEN 21 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'99') THEN 77 
                                                                                                                                                        ELSE 21 END  
                                         WHEN PWE_SCNO = N'574' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'BI') THEN 43 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'D1') THEN 46 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'D2') THEN 47 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'E1') THEN 141 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'EM') THEN 40 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'EP') THEN 45 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'IQ') THEN 43 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'NC') THEN 40 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'NE') THEN 43 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'PD') THEN 38 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'PE') THEN 42 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'SF') THEN 41 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'SS') THEN 45 
                                                                                                                                                        WHEN LEFT(PWE_SCDT,2) In (N'ST') THEN 43 
                                                                                                                                                        ELSE 43 END 
                                         WHEN PWE_SCNO = N'5741' THEN 43 
                                         WHEN PWE_SCNO = N'5761' THEN CASE WHEN LEFT(PWE_SCDT,2) In (N'75') THEN 55 
                                                                          WHEN LEFT(PWE_SCDT,2) In (N'76') THEN 56 
                                                                                                                                                        ELSE 55 END 
                                         WHEN PWE_SCNO = N'741' THEN 24 
                                         WHEN r.CalendarId IS NOT NULL THEN r.CalendarId 
                                         ELSE 99999
                                         END CalendarId
                 ,4 AS [Priority]
       ,'XRYMPWER' SourceTable
                 ,a.PWE_EMFG
                 ,a.PWE_SRTP
                 ,CASE WHEN PWE_DRSN = 'EO' THEN 1 ELSE 0 END IsDelete
into #XRYMPWER
FROM PXRYDB.SKDBA.XRYMPWER a 
INNER JOIN hisdb.dbo.PROMMedicalNote ON MedicalNoteNo = left(a.PWE_PKEY,8)  --先過濾病歷號 by Leon
LEFT JOIN hisschdb.dbo.tmpEXAMRoomMapping r ON r.RoomCode= LEFT(PWE_SCDT,2)
LEFT JOIN #XRYMDVCF1 x ON a.PWE_PKEY = x.DVC_CHRT+x.DVC_RQNO --未排程排除己有檢查室的單號 by Leon 2020-10-06
WHERE a.PWE_CHD7 >= '1090501'
AND a.PWE_EXDT = ''
AND LEN(rtrim(PWE_SCDT)) = 2
AND x.DVC_CHRT IS NULL

IF object_id('tempdb..#XRYMMWER') IS NOT NULL DROP TABLE #XRYMMWER

SELECT left(a.MWE_PKEY,8) MWE_CHRT,SUBSTRING(a.MWE_PKEY,9,9) MWE_RQNO, LEFT(MWE_SCDT,2) MWE_ROOM,a.MWE_DTNO,
       a.MWE_DEPT,a.MWE_CHD7,MWE_CHD7 MWE_TXD7,
       MWE_CHTM MWE_TXTM,a.MWE_OPNO,MWE_CHTM, 0 AS ModifyEmpId,'' AS DNL_CE_TXDT
    ,'' MWE_TEL, NULL PlanDate
              ,CASE WHEN LEFT(MWE_SCDT,2) In (N'C1', N'C2', N'C5', N'CR', N'CT', N'G1', N'M1', N'M5', N'MG', N'MR', N'PM', N'X1', N'X2', N'X3', N'X7', N'XA') AND r.CalendarId IS NOT NULL THEN r.CalendarId
                    ELSE (CASE WHEN MWE_SCNO = N'7421' THEN 110 
                                                  WHEN MWE_SCNO = N'7423' THEN 119 
                                                                      WHEN MWE_SCNO = N'7424' THEN 126 
                                                                      WHEN MWE_SCNO = N'7425' THEN 107 
                                                                      WHEN MWE_SCNO = N'7426' THEN 120 
                                                                      WHEN MWE_SCNO = N'7413' THEN 103 
                                        ELSE 110 END)  END CalendarId
                 ,5 AS [Priority]
       ,'XRYMMWER' SourceTable
                 ,a.MWE_EMFG
                 ,a.MWE_SRTP
                 ,CASE WHEN MWE_DRSN = 'EO' THEN 1 ELSE 0 END IsDelete
                 INTO #XRYMMWER
FROM PXRYDB.SKDBA.XRYMMWER a 
INNER JOIN hisdb.dbo.PROMMedicalNote ON MedicalNoteNo = left(a.MWE_PKEY,8)  --先過濾病歷號 by Leon
LEFT JOIN #XRYMDVCF1 x ON a.MWE_PKEY = x.DVC_CHRT+x.DVC_RQNO --未排程排除己有檢查室的單號 by Leon 2020-10-16
LEFT JOIN hisschdb.dbo.tmpEXAMRoomMapping r ON r.RoomCode= LEFT(MWE_SCDT,2)
WHERE a.MWE_CHD7 >= '1090501'
AND a.MWE_SCNO In (N'7413',N'7421', N'7423', N'7424', N'7425', N'7426')
AND a.MWE_EXDT = ''
AND x.DVC_CHRT IS NULL



IF OBJECT_ID('TEMPDB..#Source1') is not null DROP TABLE #Source1
SELECT A.*
--病歷+單號+檢查室 只取最後一筆 (依Priority)
,MedicalNoteId,PatientId
--,CASE WHEN LEN(ISNULL(DVC_STTM,'')) = 0 THEN '000000' WHEN LEN(ISNULL(DVC_STTM,'')) = 4  THEN RTRIM(DVC_STTM)+'00' ELSE DVC_STTM END DVC_STTM_new   
,ROW_NUMBER() OVER (PARTITION BY DVC_CHRT,DVC_RQNO,CalendarId ORDER BY [Priority] ASC, CASE WHEN  DVC_TXDT+DVC_TXTM ='' THEN DNL_CE_TXDT ELSE DVC_TXDT+DVC_TXTM END   DESC ) AS ROWNO 
INTO #Source1 
FROM (
SELECT x.DVC_CHRT,x.DVC_RQNO,x.DVC_ROOM,x.DVC_DTNO,
       x.DVC_DEP1,x.DVC_DATE,x.DVC_TXDT,
       x.DVC_TXTM,x.DVC_TXOP,DVC_STTM,x.ModifyEmpId,x.DNL_CE_TXDT
                 ,x.DVC_TEL,
                           x.PlanDate
                 ,x.CalendarId
                 ,x.[Priority]
       ,x.SourceTable
                 ,x.IsDelete
FROM #XRYMDVCF1 x
UNION ALL
--週邊血管 by Leon
SELECT OPE_CHRT,OPE_ODRN,OPE_OPRM,OPE_DOCT
      ,OPE_DEPT,OPE_DATE,OPE_TXDT
      ,OPE_TXTM,OPE_TXOP,OPE_TIME,0 AS ModifyEmpId,'' AS DNL_CE_TXDT
                ,'' AS OPE_TEL,
                           CASE WHEN ISNUMERIC(OPE_DATE) = 1 AND ISNUMERIC(OPE_TIME) = 1 
                                         THEN try_convert(datetime,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(OPE_DATE)) = 7 
                                                                     THEN OPE_DATE 
                                                                     ELSE      (CASE WHEN LEFT(OPE_DATE,1) = '0' THEN '1' ELSE '0' END) + OPE_DATE END )) ,1,3)+ 1911)+RIGHT(OPE_DATE,4) +' '+substring(OPE_TIME,1,2)+':'+substring(OPE_TIME,3,2))     
                                         WHEN ISNUMERIC(OPE_DATE) = 1 
                                         THEN try_convert(date,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(OPE_DATE)) = 7 
                                                                     THEN OPE_DATE 
                                                                    ELSE (CASE WHEN LEFT(OPE_DATE,1) = '0' THEN '1' ELSE '0' END) + OPE_DATE END )),1,3)+ 1911)+RIGHT(OPE_DATE,4))     
                                         END PlanDate
                ,CASE WHEN OPE_OPRM = '01' THEN 55 WHEN OPE_OPRM = '02' THEN 56 WHEN OPE_OPRM = '03' THEN 57 ELSE 55 END CalendarId
                ,2 AS [Priority]
      ,'OPDMOPEF' SourceTable
                ,0 IsDelete
FROM POPDDB.SKDBA.OPDMOPEF x
INNER JOIN hisdb.dbo.PROMMedicalNote ON MedicalNoteNo = OPE_CHRT --先過濾病歷號 by Leon
WHERE OPE_NOTE = '' --非取消
AND LEN(rtrim(x.OPE_ODRN)) = 9
UNION ALL

SELECT CAT_CHRT,CAT_RQNO,CAT_ROOM,CAT_DTNO
      ,CAT_DEPT,CAT_SHDT,CAT_DATE
      ,CAT_TIME,CAT_UPID,CAT_STTM,EmpId AS ModifyEmpId,'' AS DNL_CE_TXDT
                ,'' AS CAT_TEL,
                           CASE WHEN ISNUMERIC(CAT_SHDT) = 1 AND ISNUMERIC(CAT_STTM) = 1 
                                         THEN try_convert(datetime,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(CAT_SHDT)) = 7 
                                                                     THEN CAT_SHDT 
                                                                     ELSE      (CASE WHEN LEFT(CAT_SHDT,1) = '0' THEN '1' ELSE '0' END) + CAT_SHDT END )) ,1,3)+ 1911)+RIGHT(CAT_SHDT,4) +' '+substring(CAT_STTM,1,2)+':'+substring(CAT_STTM,3,2))     
                                         WHEN ISNUMERIC(CAT_SHDT) = 1 
                                         THEN try_convert(date,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(CAT_SHDT)) = 7 
                                                                     THEN CAT_SHDT 
                                                                     ELSE (CASE WHEN LEFT(CAT_SHDT,1) = '0' THEN '1' ELSE '0' END) + CAT_SHDT END )),1,3)+ 1911)+RIGHT(CAT_SHDT,4))     
                                         END PlanDate
                ,isnull(r.CalendarId,84) CalendarId
                ,3 AS [Priority]
      ,'XRYMCATF' SourceTable
                ,0 IsDelete
FROM PXRYDB.SKDBA.XRYMCATF
INNER JOIN hisdb.dbo.PROMMedicalNote ON MedicalNoteNo = CAT_CHRT --先過濾病歷號 by Leon
LEFT JOIN hisschdb.dbo.tmpEXAMRoomMapping r ON r.OldRoomCode= CAT_ROOM
LEFT JOIN hisdb.dbo.PROMEmployee ON EMPNO=CAT_UPID
UNION ALL

SELECT a.PWE_CHRT,a.PWE_RQNO,a.PWE_ROOM,a.PWE_DTNO,
       a.PWE_DEPT,a.PWE_CHD7,a.PWE_TXD7,
       a.PWE_TXTM,a.PWE_OPNO,a.PWE_CHTM,a.ModifyEmpId,a.DNL_CE_TXDT
    ,a.PWE_TEL,a.PlanDate
              ,a.CalendarId
              ,a.[Priority]
              ,a.SourceTable
              ,a.IsDelete
FROM #XRYMPWER a
WHERE LEFT(a.PWE_RQNO,3) NOT IN ('332') --移除332開頭的單 by Leon 2020-10-16
and a.PWE_EMFG <> 'Y' AND PWE_SRTP NOT IN ('E','P') --未排程-排除急作&急診&健檢 by Leon 2020-10-16

UNION ALL
SELECT a.MWE_CHRT,a.MWE_RQNO,a.MWE_ROOM,a.MWE_DTNO,
       a.MWE_DEPT,a.MWE_CHD7,a.MWE_TXD7,
       a.MWE_TXTM,a.MWE_OPNO,a.MWE_CHTM,a.ModifyEmpId,a.DNL_CE_TXDT
    ,a.MWE_TEL,a.PlanDate
              ,a.CalendarId
                 ,a.[Priority]
       ,a.SourceTable
                 ,a.IsDelete
FROM #XRYMMWER a
WHERE a.MWE_EMFG <> 'Y' AND MWE_SRTP NOT IN ('E','P') --未排程-排除急作&急診&健檢 by Leon 2020-10-27

UNION ALL

SELECT 
       I1.DNL_CHRT,I1.DNL_ODRN,I1.DNL_ROOM,SUBSTRING(LTRIM(I1.DNL_ORDR),1,4) AS DNL_DTNO,
       I1.DNL_DEPT, ISNULL(I2.DNL_DATE,I6.DNL_DATE) DNL_DATE,'' AS DNL_TXDT,
       '' AS DNL_TXTM,'' AS DNL_TXOP, ISNULL(I2.DNL_TIME,I6.DNL_TIME) DNL_TIME, 0 AS ModifyEmpId
       ,SUBSTRING(DNL_TXTM,1,8)
       +' '+CASE WHEN SUBSTRING(DNL_TXTM,9,6) BETWEEN '000000' AND '235959' THEN  SUBSTRING(DNL_TXTM,9,2)+':'+SUBSTRING(DNL_TXTM,11,2)+':'+SUBSTRING(DNL_TXTM,13,2) ELSE '00:00:00' END AS DNL_CE_TXDT
                 ,DNL_TEL,
                           CASE WHEN ISNUMERIC(ISNULL(I2.DNL_DATE,I6.DNL_DATE)) = 1 AND ISNUMERIC(ISNULL(I2.DNL_TIME,I6.DNL_TIME)) = 1 
                                         THEN try_convert(datetime,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(ISNULL(I2.DNL_DATE,I6.DNL_DATE))) = 7 
                                                                     THEN ISNULL(I2.DNL_DATE,I6.DNL_DATE) 
                                                                     ELSE      (CASE WHEN LEFT(ISNULL(I2.DNL_DATE,I6.DNL_DATE),1) = '0' THEN '1' ELSE '0' END) + ISNULL(I2.DNL_DATE,I6.DNL_DATE) END )) ,1,3)+ 1911)+RIGHT(ISNULL(I2.DNL_DATE,I6.DNL_DATE),4) +' '+substring(ISNULL(I2.DNL_TIME,I6.DNL_TIME),1,2)+':'+substring(ISNULL(I2.DNL_TIME,I6.DNL_TIME),3,2))     
                                         WHEN ISNUMERIC(ISNULL(I2.DNL_DATE,I6.DNL_DATE)) = 1 
                                         THEN try_convert(date,convert(varchar, substring((
                                         (CASE WHEN LEN(rtrim(ISNULL(I2.DNL_DATE,I6.DNL_DATE))) = 7 
                                                                     THEN ISNULL(I2.DNL_DATE,I6.DNL_DATE) 
                                                                     ELSE (CASE WHEN LEFT(ISNULL(I2.DNL_DATE,I6.DNL_DATE),1) = '0' THEN '1' ELSE '0' END) + ISNULL(I2.DNL_DATE,I6.DNL_DATE) END )),1,3)+ 1911)+RIGHT(ISNULL(I2.DNL_DATE,I6.DNL_DATE),4))     
                                         END PlanDate
                 ,CASE WHEN r.CalendarId IS NULL 
                       THEN (CASE WHEN I1.DNL_RITM = 'BNT' THEN 67
                                                                                   WHEN I1.DNL_RITM = 'BS' THEN  58
                                                                                   WHEN I1.DNL_RITM = 'BT' THEN 61 --CardNo:347
                                                                                   WHEN I1.DNL_RITM = 'CAT' THEN 84
                                                                                  --WHEN I1.DNL_RITM = 'CPE' THEN 
                                                                                   WHEN I1.DNL_RITM = 'CR' THEN 160
                                                                                   WHEN I1.DNL_RITM = 'CT' THEN 160
                                                                                   WHEN I1.DNL_RITM = 'EKG' THEN 23
                                                                                   WHEN I1.DNL_RITM = 'EMG' THEN 77
                                                                                   WHEN I1.DNL_RITM = 'ES' THEN 29
                                                                                   WHEN I1.DNL_RITM = 'IQ' THEN 43
                                                                                   WHEN I1.DNL_RITM = 'MON' THEN 175
                                                                                   WHEN I1.DNL_RITM = 'MR' THEN (CASE WHEN I1.DNL_SOUR = 'P' THEN 120 ELSE 120 END)
                                                                                   WHEN I1.DNL_RITM = 'NCV' THEN 45
                                                                                   WHEN I1.DNL_RITM = 'NM' THEN 63
                                                                                  --WHEN I1.DNL_RITM = 'NPE' THEN 
                                                                                   WHEN I1.DNL_RITM = 'OC' THEN 59
                                                                                   WHEN I1.DNL_RITM = 'OT' THEN 91
                                                                                  --WHEN I1.DNL_RITM = 'PC' THEN 
                                                                                   WHEN I1.DNL_RITM = 'PFT' THEN 65
                                                                                  --WHEN I1.DNL_RITM = 'PM' THEN 
                                                                                   WHEN I1.DNL_RITM = 'PT' THEN 103
                                                                                   WHEN I1.DNL_RITM = 'RF' THEN 62
                                                                                   WHEN I1.DNL_RITM = 'UD' THEN 83
                                                                                   WHEN I1.DNL_RITM = 'UDS' THEN 88
                                                                                   WHEN I1.DNL_RITM = 'US' THEN 27
                                                                                   WHEN I1.DNL_RITM = 'XA' THEN 110
                                                                                   WHEN I1.DNL_RITM = 'XB' THEN 110
                                                                                   WHEN I1.DNL_RITM = 'XC' THEN 58
                                                                                   WHEN I1.DNL_ROOM = '33' THEN 166
                                                                                   WHEN I1.DNL_ROOM = 'N1' THEN (CASE WHEN I1.DNL_SOUR = 'E' THEN 173 ELSE 24 END)
                                                                                   WHEN I1.DNL_COD1 = '7410223' THEN 63
                                                                                  ELSE 99999 END --找不到的大部份是牙科，暫不處理
                                         )
                       ELSE r.CalendarId END CalendarId
                 ,6 AS [Priority]
       ,'XRYTDNLD' SourceTable
                 ,0 IsDelete
FROM #XRYTDNLDI1 I1
INNER hash JOIN hisdb.dbo.PROMMedicalNote ON MedicalNoteNo=DNL_CHRT  --先過濾病歷號 by Leon
INNER hash JOIN POPDDB.SKDBA.OPDMXFEE  ON I1.DNL_COD1 = XFE_CODN
LEFT hash JOIN #XRYTDNLD I2 ON I1.DNL_CHRT = I2.DNL_CHRT and I1.DNL_ODRN = I2.DNL_ODRN AND I2.DNL_FUNC='I2' AND I2.ROWNO=1
LEFT hash JOIN #XRYTDNLD I6 ON I1.DNL_CHRT = I6.DNL_CHRT and I1.DNL_ODRN = I6.DNL_ODRN AND I6.DNL_FUNC='I6' AND I6.ROWNO=1
LEFT hash JOIN hisschdb.dbo.tmpEXAMRoomMapping r ON r.RoomCode= DNL_ROOM
left hash join #XRYMDVCF1 x1 ON I1.DNL_CHRT = x1.DVC_CHRT and I1.DNL_ODRN = x1.DVC_RQNO --排除己排程的清單 by Leon --2020-10-06
left hash join #XRYMPWER x2 ON I1.DNL_CHRT = x2.PWE_CHRT and I1.DNL_ODRN = x2.PWE_RQNO --排除己列入未排程的清單 by Leon --2020-10-06
LEFT HASH JOIN #XRYMMWER x3 ON I1.DNL_CHRT = x3.MWE_CHRT and I1.DNL_ODRN = x3.MWE_RQNO --排除己列入未排程的清單 by Leon --2020-10-14
left hash join #XRYTDNLD_D2 D2 ON I1.DNL_CHRT = D2.DNL_CHRT and I1.DNL_ODRN = D2.DNL_ODRN AND D2.DNL_FUNC='D2' AND D2.ROWNO=1 --排除取消排程的清單 by Leon --2020-10-16
WHERE I1.DNL_FUNC IN ('I1')
AND I1.ROWNO = 1
AND x1.DVC_CHRT IS NULL AND x2.PWE_CHRT is NULL AND x3.MWE_CHRT IS NULL and D2.DNL_CHRT is null
AND LEFT(i1.DNL_ODRN,3) NOT IN ('332') --移除332開頭的單 by Leon 2020-10-16
AND I1.DNL_SOUR NOT IN ('E','P')
) A
INNER JOIN hisdb.dbo.PROMMedicalNote ON MedicalNoteNo=DVC_CHRT --病歷號需可以串到



IF object_id('tempdb..#RESTTReservation') IS NOT NULL DROP TABLE #RESTTReservation

DECLARE @DVC_DATE char(4) = '1102'

SELECT a.ReservationId,a.RoomCode RESRoomCode,
       b.DVC_ROOM XRYRoomCode,a.DisplayName ,a.GroupDisplayName,
       a.MedicalNoteNo,a.ApplyFormNo,a.[Start],a.SourceCode,a.SourceTable,
          b.DVC_CHRT,b.DVC_RQNO,b.DVC_DATE,b.DVC_STTM,b.SourceTable XRYSourceCode,
          a.ReservationComment,[Priority]
INTO #RESTTReservation
FROM 
(
       SELECT distinct
                     --a.id, cal.CalendarCode RoomCode,calg.DisplayName CalendarGroupName,a.MedicalNoteNo,
                 a.ReservationId, isnull(cal.CalendarCode, f.OldRoomCode) RoomCode,
                           calg.DisplayName GroupDisplayName,a.MedicalNoteNo,
              a.ApplyFormNo,c.[Start],a.SourceCode,cal.DisplayName ,
               'RESTTReservation' SourceTable, a.ReservationComment
       FROM HISSCHDB.dbo.RESTReservationOrder a
       left JOIN HISSCHDB.dbo.RESTTimeslotRes b ON a.ReservationId = b.ReservationId
       left JOIN HISSCHDB.dbo.RESTTimeslot c ON b.TimeslotId = c.Id
          left JOIN HISSCHDB.dbo.RESTReservationOrderDetail d ON a.ReservationId = d.ReservationId
          left JOIN HISSCHDB.dbo.RESTReservationOrderDetailForEXA e ON d.ReservationDetailId = e.ReservationDetailId
          left JOIN HISDB.dbo.EXATOrderDetail f ON e.ExaOrderDetailId = f.ExaOrderDetailId
                LEFT JOIN HISSCHDB.dbo.PROMCalendar cal ON a.CalendarId = cal.id 
                LEFT JOIN HISSCHDB.dbo.PROMCalendarGroup calg ON cal.CalendarGroupId = calg.id

       WHERE a.MedicalOrderCode = 'EXA'
) a
FULL OUTER JOIN #Source1 b ON a.MedicalNoteNo = b.DVC_CHRT AND a.ApplyFormNo = b.DVC_RQNO


--SELECT * FROM #RESTTReservation WHERE MedicalNoteNo = '00177499'

 SELECT GroupDisplayName+' - '+DisplayName [Name], RESRoomCode,XRYRoomCode,DisplayName,GroupDisplayName,MedicalNoteNo,ApplyFormNo,Start,DVC_CHRT,DVC_RQNO,DVC_DATE,DVC_STTM 
                            FROM #RESTTReservation a
                            ORDER by XRYRoomCode, RESRoomCode, MedicalNoteNo
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
            //以排程群組-科室名稱作為檔案名稱
            var Name = migrationTableInfoList.OrderBy(p => p.Name)
                                   .Select(p => p.Name == null ? "Blank" : p.Name)
                                   .Distinct()
                                   .ToList();

            //Step 2.建立 各日期Sheet
            // var excelname = "Scheduler" + DateTime.Now.ToString("yyyyMMddhhmm") + ".xlsx";
            foreach (var date in datetime)
            {
                var excelname = new FileInfo(date.ToString("yyyyMMdd") + ".xlsx");
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var excel = new ExcelPackage(excelname))
                {
                    var importDBData = new ImportDBData();
                    importDBData.GenFirstSheet(excel, Name);
                    for (int sheetnum = 0; sheetnum <= Name.Count - 1; sheetnum++)
                    {
                        //Step 3.將對應的List 丟到各Sheet中
                        string displayname, groupdisplayname;
                        ExcelWorksheet sheet = excel.Workbook.Worksheets.Add(Name[sheetnum]);
                        var _name = Name[sheetnum].Split(" - ");
                        if (_name.Length == 1)
                        {
                            displayname = _name[0];
                            groupdisplayname = _name[0];
                        }
                        else {
                            displayname = _name[1];
                            groupdisplayname = _name[0];
                        }
                        //抽function
                        int rowIndex = 2;
                        int colIndex = 1;
                        importDBData.ImportData(dt, sheet, rowIndex, colIndex, migrationTableInfoList, date, groupdisplayname, displayname);
                    }
                    // Step 4.Export EXCEL
                    Byte[] bin = excel.GetAsByteArray();
                    File.WriteAllBytes(fliepath.ToString() + @"\" + excelname, bin);

                }
            }
        }

        #region --DBdata getset
        public class DBData
        {
            [Required]
            [DisplayName("新檢查室")]
            public string RESRoomCode { get; set; }
            [DisplayName("Sheet")]
            public string Name { get; set; }
            [Required]
            [DisplayName("舊檢查室")]
            public string XRYRoomCode { get; set; }
            [Required]
            [DisplayName("檢查室名稱")]
            public string DisplayName { get; set; }
            [Required]
            [DisplayName("排程群組")]
            public string GroupDisplayName { get; set; }
            [Required]
            [DisplayName("病歷號")]
            public string MedicalNoteNo { get; set; }
            [Required]
            [DisplayName("檢查單號")]
            public string ApplyFormNo { get; set; }
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
            public void ImportData(DataTable dt, ExcelWorksheet sheet, int rowIndex, int colIndex, List<DBData> dblist, DateTime date, string groupdisplayname, string displayname)
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
                for (int columnNameIndex = 1; columnNameIndex <= _dt.Columns.Count - 1; columnNameIndex++)
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
                    if (displayname == (dbdata.DisplayName == null ? "Blank" : dbdata.DisplayName) && groupdisplayname == (dbdata.GroupDisplayName == null ? "Blank" : dbdata.GroupDisplayName) && date.ToString("yyyy-MM-dd") == (dbdata.Start != DateTime.MinValue ? dbdata.Start.ToString("yyyy-MM-dd") : dbdata.PlanDate.ToString("yyyy-MM-dd")))
                    {
                        _rowIndex++;
                        _colIndex = 1;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.RESRoomCode;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.XRYRoomCode;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DisplayName;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.GroupDisplayName;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.MedicalNoteNo;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.ApplyFormNo;
                        _sheet.Cells[_rowIndex, _colIndex].Value = dbdata.Start;
                        _sheet.Cells[_rowIndex, _colIndex].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                        _sheet.Cells[_rowIndex, _colIndex++].Style.Numberformat.Format = "yyyy/MM/dd HH:mm:ss";
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_CHRT;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_RQNO;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_DATE;
                        _sheet.Cells[_rowIndex, _colIndex++].Value = dbdata.DVC_STTM;
                        if (dbdata.MedicalNoteNo == (temp_MedicalNoteNo == null ? string.Empty : temp_MedicalNoteNo) && dbdata.ApplyFormNo == (temp_ExaRequestNo == null ? string.Empty : temp_ExaRequestNo) && dbdata.DVC_CHRT == (temp_DVC_CHRT == null ? string.Empty : temp_DVC_CHRT))
                        {
                            _sheet.Cells[_rowIndex--, _colIndex].Value = "v";
                            _sheet.Cells[_rowIndex++, _colIndex].Value = "v";
                        }
                        temp_MedicalNoteNo = dbdata.MedicalNoteNo;
                        temp_ExaRequestNo = dbdata.ApplyFormNo;
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
