using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;

using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using iText.IO.Util;
using iText.Layout;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser;

namespace easyPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            //string ini파일 = args[0];
            string cfg파일 = @"c:\Program Files\Quadient\easyPDF\EasyPDF.cfg";
            if (string.IsNullOrEmpty(cfg파일)) {
                Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | EasyPDF 설정파일(*.cfg) 경로를 입력하세요", DateTime.Now));
                return;
            }
            CFG설정파일(cfg파일);
        }

        static void CFG설정파일(string cfg파일)
        {
            if (new System.IO.FileInfo(cfg파일).Exists == false) {
                Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | {1} 해당경로에 파일이 존재하지 않습니다.", DateTime.Now, cfg파일));
                return;
            }
            string[] data = System.IO.File.ReadAllLines(cfg파일, Encoding.GetEncoding(949));
            PDFSP pdfsp = new PDFSP(data);
            pdfsp.PDF설정();
        }
    }
    class PDFSP
    {
        private string[] 옵션;
        private WriterProperties setPassword;
        private DataTable 레코드info = new DataTable();

        public PDFSP(string[] 옵션)
        {
            this.옵션 = 옵션;
        }

        private DataTable 테이블정의()
        {
            DataTable tmp레코드info = new DataTable("레코드");
            string[] tmp필드 = new string[] {"일련번호", "시작페이지", "종료페이지", "패스워드", "파일명", "채널구분",
                                             "그룹01", "그룹02", "그룹03", "그룹04", "그룹05", "그룹06", "그룹07", "그룹08", "그룹09", "그룹10", "infoid", "postid", "출력여부" };
            foreach (string 필드 in tmp필드) {
                if (필드 == "시작페이지" || 필드 == "종료페이지") {
                    tmp레코드info.Columns.Add(new DataColumn(필드, typeof(Int32)));
                } else {
                    tmp레코드info.Columns.Add(new DataColumn(필드, typeof(string)));
                }
            }
            return tmp레코드info;
        }

        private string 매개변수가져오기(string 매개변수)
        {
            매개변수 = Array.Find(옵션, result => result.StartsWith(매개변수, StringComparison.OrdinalIgnoreCase)); // 대소문자 구분하지 않음
            매개변수 = (매개변수 == null) ? "" : 매개변수.Substring(매개변수.IndexOf(':') + 1);
            return 매개변수;
        }
        public void PDF설정()
        {
            // EasyPDF.cfg 파일 읽고 매개변수 Setting
            bool 확인 = false;
            try {

                // pdf입력경로(기본)
                string pdf기본경로 = 매개변수가져오기("pdf입력경로");
                if (pdf기본경로 == "") { Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | pdf입력경로 오류 ", DateTime.Now)); return; }

                // pdf출력경로
                string pdf출력경로 = "";

                // pdf입력경로(상세)
                string pdf입력경로;

                // 선택옵션 : 필요에 의해 선택 할 수 있다. ( EasyPDF.cfg에 'Y'값으로 셋팅 후 사용가능 )
                List<string> 그룹 = new List<string>();
                for (int i = 1; i <= 10; i++) {
                    그룹.Add(매개변수가져오기(string.Format("그룹{0:00}", i))); // 그룹별분리01, 그룹별분리02, 그룹별분리03 .... 10
                }

                string 개발계컴퓨터 = 매개변수가져오기("개발계컴퓨터").ToUpper();
                string 현재컴퓨터 = SystemInformation.ComputerName.ToUpper();
                string 파일이동개발 = 매개변수가져오기("파일이동개발"); // 개발계주소
                string 파일이동운영 = 매개변수가져오기("파일이동운영"); // 운영계주소

                string 개발운영주소 = "";
                // 현재컴퓨터가 개발계 인지 운영계인지 판단하여 개별파일을 이동할 주소를 정한다.
                if (현재컴퓨터 == 개발계컴퓨터) {
                    개발운영주소 = 파일이동개발;
                } else {
                    개발운영주소 = 파일이동운영;
                }

                string infoid = "";

                Console.Clear(); //화면지우기
                int 화면cnt = 0;

                string[] 분리경로 = new string[] { "M", "E" };
                for (int i = 0; i < 분리경로.Length; i++){
                    pdf입력경로 = pdf기본경로 + 분리경로[i] + "\\";
                    pdf출력경로 = pdf입력경로;
                    if (new System.IO.DirectoryInfo(pdf입력경로).Exists == false) {
                        continue;
                    }

                    System.IO.DirectoryInfo 파일목록 = new System.IO.DirectoryInfo(pdf입력경로);
                    bool inf존재여부 = false;

                    foreach (System.IO.FileInfo 파일 in 파일목록.GetFiles("*.inf")) {
                        string info파일 = 파일.FullName;
                        string pdf파일 = 파일.FullName.Replace(".inf", ".pdf");
                        화면cnt++;

                        PDFsplitter(pdf파일, pdf출력경로, info파일, 화면cnt, out infoid, out 확인);               
                        if (확인 == false) return;

                        inf존재여부 = true; // 개별파일 분류 
                        
                    }
                    // 파일복사 후 삭제
                    
                    if (inf존재여부) {
                        try {
                            string 파일이동주소 = 개발운영주소 + string.Format(@"{0}\{1:yyyyMMdd}\", 분리경로[i], DateTime.Now);
                            if (new System.IO.DirectoryInfo(파일이동주소).Exists == false) {
                                System.IO.Directory.CreateDirectory(파일이동주소);
                            }

                            파일이동주소 += string.Format(@"{0}\", infoid);
                            if (new System.IO.DirectoryInfo(파일이동주소).Exists == false) {
                                System.IO.Directory.CreateDirectory(파일이동주소);
                            }

                            Console.ForegroundColor = ConsoleColor.Red;
                            
                            int 복사파일수 = new DirectoryInfo(pdf출력경로).GetFiles("form_*.pdf").Length;
                            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | 개별파일 {1:#,#}개 복사시작 ", DateTime.Now, 복사파일수));

                          
                            Console.WriteLine("***복사중** : ");
                            int 복사cnt = 0;


                            int 자리수 = 복사파일수.ToString().Length;
                            int 단위 = 1;

                            if (자리수 == 1) {
                                단위 = 1;
                            } else {
                                for (int ii = 1; ii < 자리수; ii++)
                                {
                                    단위 = 단위 * 10;
                                }
                            }

                            //form_
                            foreach (System.IO.FileInfo 파일 in new DirectoryInfo(pdf출력경로).GetFiles("form_*.pdf")){    
                                new System.IO.FileInfo(파일.FullName).CopyTo(string.Format("{0}{1}", 파일이동주소, 파일.Name), true);
                                복사cnt++;
                                
                                if ((복사cnt % 단위  == 0) || 복사cnt ==  복사파일수) {
                                    int 백분율 = 복사cnt / 복사파일수 * 100;
                                    Console.WriteLine(string.Format("{0:0000000} / {1:0000000} ...Copying  {2:000}%  ", 복사cnt, 복사파일수, 백분율));
                                    Application.DoEvents();
                                }

                                
                            }
                            Console.WriteLine();
                            Console.WriteLine();
                            Application.DoEvents();

                            System.Threading.Thread.Sleep(10000); //10초 삭제



                            // 파일삭제
                            foreach (System.IO.FileInfo 파일 in new DirectoryInfo(pdf출력경로).GetFiles("*.*")) {
                                new System.IO.FileInfo(파일.FullName).Delete();
                            }
                            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | === 개별파일이동 완료 === ", DateTime.Now));

                            

                        } catch (Exception e) {
                            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | 파일이동오류 - {1}", DateTime.Now, e.Message));
                            return;
                        }

                    }
                    화면cnt++;
                }
                
            } catch (Exception e) {
                Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | {1}", DateTime.Now, e.Message));
                return;
            }
            확인 = true;
        }

        private void PDFsplitter(string pdf파일, string pdf출력경로, string info파일 , int 화면cnt, out string infoid , out bool 확인)
        {
            확인 = false;
            infoid = "";
            레코드info = 테이블정의();
            string[] data = File.ReadAllLines(info파일, Encoding.GetEncoding(949));

            for (int i = 0; i < data.Length; i++) {
                string[] 분리 = data[i].Split('|');
                /*
                row = 레코드info.NewRow();
                레코드info.Rows.Add(row);
                for (int ii = 0; ii < 분리.Length; ii++)
                    row.SetField(ii, 분리[ii]);
                    */
                레코드info.Rows.Add(new object[] {
                    분리[0],
                    Convert.ToInt32(분리[1]),
                    Convert.ToInt32(분리[2]),
                    분리[3],
                    분리[4],
                    분리[5],
                    분리[6],
                    분리[7],
                    분리[8],
                    분리[9],
                    분리[10],
                    분리[11],
                    분리[12],
                    분리[13],
                    분리[14],
                    분리[15],
                    분리[16],   //infoid
                    분리[17],   //postid
                    분리[18]

                });
                infoid = 분리[16];
            }

            PDF개별파일저장(pdf파일, pdf출력경로, 레코드info, 화면cnt, out 확인);
            if (확인 == false) return;

            확인 = true;
        }

        private void PDF개별파일저장(string pdf파일, string pdf출력경로, DataTable 레코드info, int 화면cnt, out bool 확인)
        {
            확인 = false;

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | 개별PDF로 저장 시작 - {1}", DateTime.Now, pdf파일));

            
            Console.WriteLine("***진행률** : ");
            
            FileStream fsr = new FileStream(pdf파일, FileMode.Open, FileAccess.Read, FileShare.Read);
            PdfDocument pdfDocR = new PdfDocument(new PdfReader(fsr));
            Document docR = new Document(pdfDocR);

            PageRange pageRange = null;
            int S페이지, E페이지;
            string pdf저장 = "";
            string userPass = "";
            string owerPass = "XXXXXXXXXXXXXX";

            DataRow[] selectRow = 레코드info.Select("출력여부 = '출력' ");
            int 레코드갯수 = selectRow.Length;

            double 진행률 = 0;
            double cnt = 0;
            double tot = 레코드갯수;
            int 진행막대 = 0;
            int 증가막대 = 0;
            int 현재막대 = 0;

            string 파일명 = "";
            string 시작비교값 = "";
            string 종료비교값 = "";
            PdfPage 비교page;
            Rectangle 범위 = new iText.Kernel.Geom.Rectangle(MMtoPoint(5), MMtoPoint(0), MMtoPoint(30), MMtoPoint(5));  // x,y,w,h - 왼쪽아래 기준

            Stopwatch sw = new Stopwatch();
            sw.Start();
            for (int i = 0; i < 레코드갯수; i++)
            {
                파일명 = selectRow[i]["파일명"].ToString();

                S페이지 = Convert.ToInt32(selectRow[i]["시작페이지"].ToString());
                E페이지 = Convert.ToInt32(selectRow[i]["종료페이지"].ToString());
                userPass = selectRow[i]["패스워드"].ToString();
                pageRange = new PageRange().AddPageSequence(S페이지, E페이지);
                pdf저장 = string.Format("{0}\\{1}", pdf출력경로, 파일명);   // ex) \tmp\M\abc12345.pdf


                if (i % 2 == 0) {
                    비교page = pdfDocR.GetPage(S페이지);
                    유효문자파싱(범위, 비교page, out 시작비교값);

                    if (S페이지 == E페이지) {
                        종료비교값 = 시작비교값;
                    }
                    else {
                        비교page = pdfDocR.GetPage(E페이지);
                        유효문자파싱(범위, 비교page, out 종료비교값);
                    }

                    bool postid비교 = selectRow[i]["postid"].ToString() == 시작비교값 ? (시작비교값 == 종료비교값 ? true : false) : false;
                    if ( postid비교 == false)
                    {
                        Console.WriteLine();
                        Console.WriteLine("========================================== [ 실행불가 ] ==========================================");
                        Console.WriteLine(string.Format("Error - POSTID 불일치 : {0}, 파일명 - {1}", selectRow[i]["postid"].ToString(), pdf파일));
                        Console.WriteLine(string.Format("Error -  시작페이지 : {0}, 종료페이지 : {1}", S페이지, E페이지));
                        Console.WriteLine("==================================================================================================");

                        return;
                    }

                }
                // 개별PDF 저장
                setPassword = null;
                if (userPass != "")
                    setPW(out setPassword, userPass, owerPass);

                IList<PdfDocument> splitDocuments = new PDF분리(pdfDocR, pdf저장, setPassword).ExtractPageRanges(JavaUtil.ArraysAsList(pageRange));
                foreach (PdfDocument pdfDocument in splitDocuments) {
                    pdfDocument.Close();
                }


                cnt++;
                진행률 = cnt / tot * 100;
                진행막대 = (int)(진행률 / 5);
                증가막대 = 진행막대 - 현재막대;
                현재막대 = 진행막대;
                Console.ForegroundColor = ConsoleColor.Blue;

                Console.WriteLine(string.Format("{0:00.00}% ({1} / {2})", 진행률, cnt, tot));


                for (int j = 1; j <= 증가막대; j++) {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(string.Format("■", i));
                }
            }
            sw.Stop();
            double 걸린시간 = sw.ElapsedMilliseconds / 1000.0 / 60.0;

            Console.WriteLine( string.Format("***처리시간** : {0:0.00}분", 걸린시간));
            
            docR.Close(); pdfDocR.Close(); fsr.Close();
            
            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | 개별PDF로 저장 완료", DateTime.Now));
            Console.WriteLine();
            확인 = true;
        }

        private void 유효문자파싱(Rectangle 범위, PdfPage page, out string 비교값)
        {
            //try
            //{
            FilteredTextEventListener filterListener = new FilteredTextEventListener(new LocationTextExtractionStrategy(), new iText.Kernel.Pdf.Canvas.Parser.Filter.TextRegionEventFilter(범위));
            비교값 = PdfTextExtractor.GetTextFromPage(page, filterListener);
            비교값 = Encoding.GetEncoding(949).GetString(Encoding.Default.GetBytes(비교값)).ToString().Trim();
            //} catch (Exception) {
            //    비교값 = "";
            //}

        }

        private void PDF그룹별파일저장(string pdf파일, string pdf출력경로, DataTable 레코드info, string 정렬필드, out bool 확인)
        {
            확인 = false;
            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | PDF그룹별 저장 시작 ", DateTime.Now));

            string 정렬 = string.Format("{0},시작페이지 ASC", 정렬필드);
            DataRow[] selectRow = 레코드info.Select("출력여부 = '출력' or 출력여부 = '라벨'", 정렬);

            if (string.IsNullOrEmpty(selectRow[0][정렬필드].ToString())) {
                Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | PDF그룹 : {1}, *.inf파일에 값이 존재하지 않습니다.   ", DateTime.Now, 정렬필드));
                확인 = true;
                return;
            }

            FileStream fsr = new FileStream(pdf파일, FileMode.Open, FileAccess.Read, FileShare.Read);
            PdfDocument pdfDocR = new PdfDocument(new PdfReader(fsr));
            Document docR = new Document(pdfDocR);


            string 기본키, 비교키;
            비교키 = selectRow[0][정렬필드].ToString();
            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | PDF그룹 : {1} ", DateTime.Now, 비교키));
            FileStream fsw = new FileStream(string.Format("{0}", pdf출력경로 + 비교키), FileMode.Create, FileAccess.Write, FileShare.Write);
            PdfDocument pdfDocW = new PdfDocument(new PdfWriter(fsw));
            Document docW = new Document(pdfDocW);
            PdfMerger merger = new PdfMerger(pdfDocW).SetCloseSourceDocuments(false);

            int 레코드갯수 = selectRow.Length;
            int 시작page, 종료page;
            for (int i = 0; i < 레코드갯수; i++) {
                기본키 = selectRow[i][정렬필드].ToString();

                if (기본키 != 비교키) {
                    비교키 = 기본키;
                    merger.Close(); docW.Close(); pdfDocW.Close(); fsw.Close();

                    Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | PDF그룹 : {1} ", DateTime.Now, 비교키));
                    fsw = new FileStream(string.Format("{0}", pdf출력경로 + 비교키), FileMode.Create, FileAccess.Write, FileShare.Write);
                    pdfDocW = new PdfDocument(new PdfWriter(fsw));
                    docW = new Document(pdfDocW);
                    merger = new PdfMerger(pdfDocW).SetCloseSourceDocuments(false);
                }

                시작page = Convert.ToInt32(selectRow[i]["시작페이지"].ToString());
                종료page = Convert.ToInt32(selectRow[i]["종료페이지"].ToString());
                merger.Merge(pdfDocR, 시작page, 종료page);
            }
            merger.Close(); docW.Close(); pdfDocW.Close(); fsw.Close();
            Console.WriteLine(string.Format("{0:yy.MM.dd HH:mm:ss} | PDF그룹별 저장 완료 ", DateTime.Now));

            확인 = true;
        }

       

        private float MMtoPoint(double x)
        {
            float tmpx;
            tmpx = (float)(x * 2.83465);
            return tmpx;
        }
        private void setPW(out WriterProperties setPassword, string userPass, string owerPass)
        {
            setPassword = null;
            // pdf password 설정
            
            byte[] byteuserPass = Encoding.GetEncoding(949).GetBytes(userPass);
            byte[] byteowerPass = Encoding.GetEncoding(949).GetBytes(owerPass);

            setPassword = new WriterProperties().SetStandardEncryption(byteuserPass, byteowerPass, EncryptionConstants.ALLOW_PRINTING,
                                 EncryptionConstants.ENCRYPTION_AES_256 | EncryptionConstants.DO_NOT_ENCRYPT_METADATA);
        }

        private sealed class PDF분리 : PdfSplitter
        {
            private string 출력파일명;
            private WriterProperties setPassword;
            public PDF분리(PdfDocument baseArg1, string 출력파일명, WriterProperties setPassword) : base(baseArg1)
            {
                this.출력파일명 = 출력파일명;
                this.setPassword = setPassword;
            }

            protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
            {
                try {
                    if (setPassword != null) {
                        return new PdfWriter(출력파일명, setPassword);

                    // password 없음
                    } else { 
                        return new PdfWriter(출력파일명);
                    }
                } catch (FileNotFoundException) {
                    throw new Exception();
                }
            }

        }


    }

}
