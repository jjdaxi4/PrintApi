using MesSystem;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Ghostscript.NET.Processor;
using System.IO;
using Serilog;

namespace TestProject
{

    

    /// <summary>
    /// 메인 폼
    /// </summary>
    public partial class MainForm : Form
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////// Field
        ////////////////////////////////////////////////////////////////////////////////////////// Private

        #region Field

        /// <summary>
        /// 서버
        /// </summary>
        private WebServer server = null;

        private HttpListener httpListener;

        public object ThreadPool { get; private set; }

        #endregion




        //////////////////////////////////////////////////////////////////////////////////////////////////// Constructor
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 생성자 - MainForm()

        /// <summary>
        /// 생성자
        /// </summary>
        public MainForm()
        {
            InitializeComponent();


            this.server = new WebServer();

            this.server.AddBindingAddress("http://localhost:9999/"); // 포트 번호까지 반드시 설정합니다.

            this.server.RootPath = "c:\\wwwroot";

            this.server.ActionRequested += server_ActionRequested;


            FormClosing += Form_FormClosing;

            this.startButton.Click += startButton_Click;


            
        }

        #endregion


        //////////////////////////////////////////////////////////////////////////////////////////////////// Method
        ////////////////////////////////////////////////////////////////////////////////////////// Private

        private void Main_Load(object sender, System.EventArgs e)
        {
            //자동실행
            this.Activated += startButton_Click;

            
        }


        #region 폼을 닫을려는 경우 처리하기 - Form_FormClosing(sender, e)

        /// <summary>
        /// 폼을 닫을려는 경우 처리하기
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(this.server.IsRunning)
            {
                this.server.Stop();

                this.server = null;
            }
        }

        #endregion
        #region Start 버튼 클릭시 처리하기 - startButton_Click(sender, e)

        /// <summary>
        /// Start 버튼 클릭시 처리하기
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void startButton_Click(object sender, EventArgs e)
        {
            if(this.server.IsRunning)
            {
                this.server.Stop();

                this.startButton.Text = "Start";

                this.richTextBox1.Text = "제브라 프린터 시스템이 종료되었습니다.";

                this.richTextBox2.Text = "프린터 출력 대기상태가 아닙니다. Start를 눌러 실행시켜주세요.";
            }
            else
            {


                this.server.Start();
                
                serverInit();

                this.startButton.Text = "Stop";

                this.richTextBox1.Text = "제브라 프린터 시스템이 구동중입니다.";

                this.richTextBox2.Text = "프린터 출력 대기중입니다.";

            }
        }

        #endregion

        private void serverInit()
        {
            if (httpListener == null)
            {
                //API 세팅 주소
                httpListener = new HttpListener();
                httpListener.Prefixes.Add(string.Format("http://localhost:8585/"));                
                //httpListener.Prefixes.Add(string.Format("http://+:80/"));
                serverStart();
            }
        }

        private void serverStart()
        {

            if (!httpListener.IsListening)
            {
                httpListener.Start();                

                Task.Factory.StartNew(() =>
                {
                    while (httpListener != null)
                    {

                        HttpListenerContext ctx = this.httpListener.GetContext();

                        //string methodName = ctx.Request.Url.Segments[1].Replace("/", "");
                        //string strParams = System.Net.WebUtility.UrlEncode(ctx.Request.Url.Segments[2].Replace("/", ""));

                        string strParams = ctx.Request.Url.LocalPath.Replace("/param=","");

                        PrintDialog pd = new PrintDialog();
                        pd.PrinterSettings = new PrinterSettings();

                        PrinterSettings oPrinterSettings = new PrinterSettings();

                        string printerName = oPrinterSettings.PrinterName;

                        //USER 프린터일 경우
                        if (strParams.Contains("pdf"))
                        {
                            string inputFile = @"D:\D365_PRINT\PDF\" + strParams;

                            using (GhostscriptProcessor processor = new GhostscriptProcessor())
                            {
                                List<string> switches = new List<string>();
                                switches.Add("-empty");
                                switches.Add("-dPrinted");
                                switches.Add("-dBATCH");
                                switches.Add("-dNOPAUSE");
                                //switches.Add("-dQUIET");
                                switches.Add("-dNOSAFER");
                                switches.Add("-dNoCancel");
                                switches.Add("-dNumCopies=1");
                                switches.Add("-dQueryUser=3");
                                switches.Add("-dORIENT1=false");
                                //switches.Add("-sPAPERSIZE=a4");

                                //switches.Add("-dNORANGEPAGESIZE");
                                switches.Add("-sDEVICE=mswinpr2");
                                switches.Add("-sOutputFile=%printer%" + printerName);
                                switches.Add("-f");
                                switches.Add(inputFile);


                                try
                                {
                                    processor.StartProcessing(switches.ToArray(), null);
                                    CreateFileSerilog("ZPL|" + printerName + "|" + inputFile);
                                }
                                catch (Exception e) {
                                    CreateFileSerilog("ZPL_ERROR|" + e.Message.ToString());
                                }
                                finally
                                {

                                    if (richTextBox2.InvokeRequired)
                                        richTextBox2.Invoke(new MethodInvoker(delegate { richTextBox2.Text = "출력 성공"; }));
                                }

                            }


                            //60일 지난 파일 삭제
                            try
                            {
                                Delete_File("PDF");                                
                            }
                            catch (Exception e)
                            {
                                CreateFileSerilog("ZPL_FILEDEL_ERROR|" + e.Message.ToString());
                            }


                        }
                        else { //제브라 프린터

                            dynamic dynJson = JsonConvert.DeserializeObject(strParams);
                            string result = string.Empty;
                            string barcode = string.Empty;
                            foreach (var item in dynJson)
                            {

                                result += string.Format("rawurl = {0}\r\n", item.TEXT1);
                                result += string.Format("rawurl = {0}\r\n", item.TEXT2);

                                barcode += this.getbarcode(Convert.ToString(item.TEXT1), Convert.ToString(item.TEXT2));
                            }


                            if (richTextBox1.InvokeRequired)
                                richTextBox1.Invoke(new MethodInvoker(delegate { richTextBox1.Text = result; }));
                            else
                                richTextBox1.Text = result;


                            try
                            {
                                RawPrinterHelper.SendStringToPrinter(pd.PrinterSettings.PrinterName, barcode);

                                CreateFileSerilog("PRINT|" + pd.PrinterSettings.PrinterName + "|" + barcode);

                            }
                            catch (Exception e)
                            {

                                if (richTextBox2.InvokeRequired)
                                    richTextBox2.Invoke(new MethodInvoker(delegate { richTextBox2.Text = e.Message.ToString(); }));                               
                                
                                CreateFileSerilog("PRINT_ERROR|" + e.Message.ToString());
                            }
                            finally
                            {

                                if (richTextBox2.InvokeRequired)
                                    richTextBox2.Invoke(new MethodInvoker(delegate { richTextBox2.Text = "출력 성공"; }));     
                                
                            }


                            ctx.Response.Close();

                            //60일 지난 로그 파일 삭제
                            Delete_File("LOG");

                        }

                    }
                });

            }
        }

        

        /// <summary>바코드 출력</summary>
        private string getbarcode(string text1, string text2)
        {
            string BarCode = string.Empty;            

            BarCode = "^XA";
            BarCode = BarCode + "^FX BOX DRAW";
            BarCode = BarCode + "^FO10,40^GB530,290,3^FS";
            BarCode = BarCode + "^FO10,40^GB530,335,3^FS";

            BarCode = BarCode + "^FX ITEMID AND Sn";
            BarCode = BarCode + "^CFA,30 ";

            BarCode = BarCode + "^FX ITEM CD BARCODE (VALUE)";
            BarCode = BarCode + "^BY3,2,90";
            BarCode = BarCode + "^FO20,50^BC,126,N^" + text1 + "^FS";

            BarCode = BarCode + "^FX ITEM CD TEXT (FIX)";
            BarCode = BarCode + "^FO20,150^FDITEM CD :^FS";

            BarCode = BarCode + "^FX SN BARCODE(VALUE)";
            BarCode = BarCode + "^BY3,2,90";
            BarCode = BarCode + "^FO20,190^BC^FD" + text2 + "^FS";


            BarCode = BarCode + "^FX SN BARCODE(VALUE)";
            BarCode = BarCode + "^BY3,2,90";
            BarCode = BarCode + "^FO20,190^BC^" + text2 + "^FS";

            BarCode = BarCode + "^FX SN TEXT (FIX)";
            BarCode = BarCode + "^FO20,290^FDSN :^FS ";

            BarCode = BarCode + "^FX SN TEXT (VALUE)";
            BarCode = BarCode + "^FO100,290^FD" + text2 + "^FS";

            BarCode = BarCode + "^CF0,30,30";
            BarCode = BarCode + "^FO70,340^FDMANUFACTURED BY TYM. KOREA^FS";
            BarCode = BarCode + "^XZ";

            return BarCode;
            
        }

        #region 60일 지난  파일 삭제

        private void Delete_File(string path) {

            string deletePath = @"D:\D365_PRINT/" + path;

            (from f in new DirectoryInfo(deletePath).GetFiles()
             where f.CreationTime < DateTime.Now.Subtract(TimeSpan.FromDays(60))
             select f
            ).ToList()
                .ForEach(f => f.Delete());
        }

        #endregion


        #region Log 기록
        //Log 파일에 기록
        private static void CreateFileSerilog(string logTxt)
        {
            Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.Console()
            .WriteTo.File(@"D:\D365_PRINT/LOG/log.txt",
                rollingInterval: RollingInterval.Day,
                rollOnFileSizeLimit: true)
            .CreateLogger();

            Log.Information(logTxt);
 

            Log.CloseAndFlush();
        }
        #endregion

        #region 서버 액션 요청시 처리하기 - server_ActionRequested(sender, e)

        /// <summary>
        /// 서버 액션 요청시 처리하기
        /// </summary>
        /// <param name="sender">이벤트 발생자</param>
        /// <param name="e">이벤트 인자</param>
        private void server_ActionRequested(object sender, ActionRequestedEventArgs e)
        {
            e.Server.WriteDefaultAction(e.Context);
        }

        #endregion
    }
}