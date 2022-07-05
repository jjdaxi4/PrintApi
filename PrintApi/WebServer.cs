using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace TestProject
{
    /// <summary>
    /// 웹 서버
    /// </summary>
    public class WebServer : IDisposable
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////// Event
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 액션 요청시 이벤트 - ActionRequested

        /// <summary>
        /// 액션 요청시 이벤트
        /// </summary>
        public event EventHandler<ActionRequestedEventArgs> ActionRequested;

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Field
        ////////////////////////////////////////////////////////////////////////////////////////// Static
        //////////////////////////////////////////////////////////////////////////////// Private

        #region Field

        /// <summary>
        /// MIME 타입 딕셔너리
        /// </summary>
        private static Dictionary<string, string> _mimeTypeDictionary = null;

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////// Instance
        //////////////////////////////////////////////////////////////////////////////// Private

        #region Field

        /// <summary>
        /// 청취자
        /// </summary>
        private HttpListener listener = null;

        /// <summary>
        /// 바인딩 주소 리스트
        /// </summary>
        private List<string> bindingAddressList = new List<string>();

        /// <summary>
        /// 루트 경로
        /// </summary>
        private string rootPath = null;

        /// <summary>
        /// 청취 스레드
        /// </summary>
        private Thread listenThread = null;

        /// <summary>
        /// 응답 스레드 리스트
        /// </summary>
        private List<Thread> responseThreadList = null;

        /// <summary>
        /// 실행 여부
        /// </summary>
        private bool isRunning = false;

        /// <summary>
        /// 해제 여부
        /// </summary>
        private bool isDisposed = false;

        /// <summary>
        /// 버퍼 크기
        /// </summary>
        private const int BUFFER_SIZE = 4096;

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Property
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 루트 경로 - RootPath

        /// <summary>
        /// 루트 경로
        /// </summary>
        public string RootPath
        {
            get
            {
                return this.rootPath;
            }
            set
            {
                if(this.isRunning)
                {
                    return;
                }

                this.rootPath = value;
            }
        }

        #endregion
        #region 실행 여부 - IsRunning

        /// <summary>
        /// 실행 여부
        /// </summary>
        public bool IsRunning
        {
            get
            {
                return this.isRunning;
            }
        }

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Constructor
        ////////////////////////////////////////////////////////////////////////////////////////// Static

        #region 생성자 - WebServer()

        /// <summary>
        /// 생성자
        /// </summary>
        static WebServer()
        {
            _mimeTypeDictionary = new Dictionary<string, string>();

            _mimeTypeDictionary.Add(".js"  , "application/js"          );
            _mimeTypeDictionary.Add(".json", "application/json"        );
            _mimeTypeDictionary.Add(".html", "text/html; charset=utf-8");
            _mimeTypeDictionary.Add(".css" , "text/css; charset=utf-8" );
            _mimeTypeDictionary.Add(".text", "text/text; charset=utf-8");
            _mimeTypeDictionary.Add(".jpg" , "image/jpeg"              );
            _mimeTypeDictionary.Add(".png" , "image/png"               );
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////// Instance
        //////////////////////////////////////////////////////////////////////////////// Public

        #region 생성자 - WebServer()

        /// <summary>
        /// 생성자
        /// </summary>
        public WebServer()
        {
        }

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Method
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 바인딩 주소 포함 여부 구하기 - ContainsBindingAddress(bindingAddress)

        /// <summary>
        /// 바인딩 주소 포함 여부 구하기
        /// </summary>
        /// <param name="bindingAddress">바인딩 주소</param>
        /// <returns>바인딩 주소 포함 여부</returns>
        public bool ContainsBindingAddress(string bindingAddress)
        {
            return this.bindingAddressList.Contains(bindingAddress);
        }

        #endregion
        #region 바인딩 주소 추가하기 - AddBindingAddress(bindingAddress)

        /// <summary>
        /// 바인딩 주소 추가하기
        /// </summary>
        /// <param name="bindingAddress"></param>
        public void AddBindingAddress(string bindingAddress)
        {
            if(this.isRunning)
            {
                return;
            }

            if(ContainsBindingAddress(bindingAddress))
            {
                return;
            }

            this.bindingAddressList.Add(bindingAddress);
        }

        #endregion
        #region 바인딩 주소 제거하기 - RemoveBindingAddress(bindingAddress)

        /// <summary>
        /// 바인딩 주소 제거하기
        /// </summary>
        /// <param name="bindingAddress">바인딩 주소</param>
        public void RemoveBindingAddress(string bindingAddress)
        {
            if(this.isRunning)
            {
                return;
            }

            if(!ContainsBindingAddress(bindingAddress))
            {
                return;
            }

            this.bindingAddressList.Remove(bindingAddress);
        }

        #endregion
        #region 바인딩 주소 리스트 지우기 - ClearBindingAddressList()

        /// <summary>
        /// 바인딩 주소 리스트 지우기
        /// </summary>
        public void ClearBindingAddressList()
        {
            if(this.isRunning)
            {
                return;
            }

            this.bindingAddressList.Clear();
        }

        #endregion

        #region 시작하기 - Start()

        /// <summary>
        /// 시작하기
        /// </summary>
        public void Start()
        {
            if(this.isRunning)
            {
                return;
            }

            if(this.bindingAddressList == null || this.bindingAddressList.Count == 0)
            {
                return;
            }

            this.listener = new HttpListener();

            foreach(string bindingAddress in this.bindingAddressList)
            {
                this.listener.Prefixes.Add(bindingAddress);
            }

            this.listener.AuthenticationSchemes = AuthenticationSchemes.Basic | AuthenticationSchemes.Anonymous;

            this.listener.Start();

            this.isRunning = true;

            if(this.responseThreadList != null)
            {
                foreach(Thread responseThread in this.responseThreadList)
                {
                    responseThread.Abort();
                }

                this.responseThreadList.Clear();
            }

            this.responseThreadList = new List<Thread>();


            this.listenThread = new Thread(new ThreadStart(Listen));
        
            this.listenThread.IsBackground = true;
        
            this.listenThread.Start();
        }

        #endregion
        #region 중단하기 - Stop()

        /// <summary>
        /// 중단하기
        /// </summary>
        public void Stop()
        {
            this.isRunning = false;
            

            if(this.listenThread != null)
            {
                this.listenThread.Abort();

                this.listenThread = null;
            }
            

            if(this.responseThreadList != null)
            {
                foreach(Thread responseThread in this.responseThreadList)
                {
                    responseThread.Abort();
                }

                this.responseThreadList.Clear();
            }


            if(this.listener != null)
            {
                this.listener.Stop();

                this.listener.Close();

                this.listener = null;
            }
        }

        #endregion

        #region POST 데이터 구하기 - GetPOSTData(request)

        /// <summary>
        /// POST 데이터 구하기
        /// </summary>
        /// <param name="request">요청</param>
        /// <returns>POST 데이터</returns>
        public string GetPOSTData(HttpListenerRequest request)
        {
            if(!request.HasEntityBody)
            {
                return null;
            }

            using(StreamReader reader = new StreamReader(request.InputStream, request.ContentEncoding))
            {
                return reader.ReadToEnd();
            }
        }

        #endregion
        #region 쿼리 문자열 구하기 - GetQueryString(request)

        /// <summary>
        /// 쿼리 문자열 구하기
        /// </summary>
        /// <param name="request">요청</param>
        /// <returns>쿼리 문자열</returns>
        public string GetQueryString(HttpListenerRequest request)
        {
            if(request.QueryString == null || request.QueryString.Count == 0)
            {
                return null;
            }

            StringBuilder stringBuilder = new StringBuilder();

            foreach(string key in request.QueryString.AllKeys)
            {
                stringBuilder.Append("&");
                stringBuilder.AppendFormat("{0}={1}", key, request.QueryString[key]);
            }

            return stringBuilder.ToString().TrimStart('&');
        }

        #endregion

        #region 파일 쓰기 - WriteFile(request, filePath)

        /// <summary>
        /// 파일 쓰기
        /// </summary>
        /// <param name="response">응답</param>
        /// <param name="filePath">파일 경로</param>
        public void WriteFile(HttpListenerResponse response, string filePath)
        {
            string mimeType = GetMIMEType(filePath);

            response.Headers.Add(HttpResponseHeader.ContentType, mimeType);

            using(BinaryReader reader = new BinaryReader(new FileStream(filePath, FileMode.Open, FileAccess.Read)))
            {
                byte[] byteArray = new byte[BUFFER_SIZE];

                int readCount;

                while((readCount = reader.Read(byteArray, 0, byteArray.Length)) > 0)
                {
                    response.OutputStream.Write(byteArray, 0, readCount);
                }
            }
        }

        #endregion
        #region 쓰기 - Write(response, message, addHeader)

        /// <summary>
        /// 쓰기
        /// </summary>
        /// <param name="response">응답</param>
        /// <param name="message">메시지</param>
        /// <param name="addHeader">헤더 추가 여부</param>
        public void Write(HttpListenerResponse response, string message, bool addHeader = false)
        {
            if(addHeader)
            {
                response.Headers.Add(HttpResponseHeader.ContentType, "text/html; charset=utf-8");
            }

            byte[] byteArray = Encoding.UTF8.GetBytes(message);

            response.OutputStream.Write(byteArray, 0, byteArray.Length);
        }

        #endregion
        #region 디폴트 액션 쓰기 - WriteDefaultAction(context)

        /// <summary>
        /// 디폴트 액션 쓰기
        /// </summary>
        /// <param name="context">컨텍스트</param>
        public void WriteDefaultAction(HttpListenerContext context)
        {
            StringBuilder stringBuilder = new StringBuilder();

            stringBuilder.AppendFormat("요청 일시 : {0}", DateTime.Now);
            stringBuilder.AppendLine("<br>");
            stringBuilder.AppendFormat("요청 URL : {0}", context.Request.Url);
            stringBuilder.AppendLine("<br>");
            stringBuilder.AppendFormat("액션명 : {0}", context.Request.RawUrl.Substring(0, context.Request.RawUrl.IndexOf("?")));
            stringBuilder.AppendLine("<br>");
            stringBuilder.AppendFormat("요청 종류 : {0}", context.Request.HttpMethod);
            stringBuilder.AppendLine("<br>");
            stringBuilder.AppendFormat("POST DATA : {0}", GetPOSTData(context.Request));
            stringBuilder.AppendLine("<br>");
            stringBuilder.AppendFormat("QUERY STRING : {0}", GetQueryString(context.Request));
            stringBuilder.AppendLine("<br>");

            Write(context.Response, stringBuilder.ToString(), true);
        }

        #endregion

        #region 리소스 해제하기 - Dispose()

        /// <summary>
        /// 리소스 해제하기
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////// Protected

        #region 리소스 해제하기 - Dispose(disposing)

        /// <summary>
        /// 리소스 해제하기
        /// </summary>
        /// <param name="disposing">해제 여부</param>
        protected virtual void Dispose(bool disposing)
        {
            if(!this.isDisposed)
            {
                if(disposing)
                {
                    Stop();
                }

                this.isDisposed = true;
            }
        }

        #endregion

        ////////////////////////////////////////////////////////////////////////////////////////// Private

        #region 물리적 경로 구하기 - GetPhysicalPath(logicalPath)

        /// <summary>
        /// 물리적 경로 구하기
        /// </summary>
        /// <param name="logicalPath">논리적 경로</param>
        /// <returns>물리적 경로</returns>
        private string GetPhysicalPath(string logicalPath)
        {
            string physicalPath = logicalPath.Replace('/', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar);

            return physicalPath;
        }

        #endregion
        #region MIME 타입 구하기 - GetMIMEType(filePath)

        /// <summary>
        /// MIME 타입 구하기
        /// </summary>
        /// <param name="filePath">파일 경로</param>
        /// <returns>MIME 타입</returns>
        private string GetMIMEType(string filePath)
        {
            string fileExtension = Path.GetExtension(filePath);

            if(_mimeTypeDictionary.ContainsKey(fileExtension))
            {
                return _mimeTypeDictionary[fileExtension];
            }
            else
            {
                return "application/octet-stream";
            }
        }

        #endregion

        #region 디폴트 요청 여부 구하기 - IsDefaultRequest(request)

        /// <summary>
        /// 디폴트 요청 여부 구하기
        /// </summary>
        /// <param name="request">요청</param>
        /// <returns>디폴트 요청 여부</returns>
        private bool IsDefaultRequest(HttpListenerRequest request)
        {
            bool isDefaultRequest = (request.Url.AbsolutePath == string.Empty || request.Url.AbsolutePath == "/");

            return isDefaultRequest;
        }

        #endregion
        #region 파일 요청 여부 구하기 - IsFileRequest(request)

        /// <summary>
        /// 파일 요청 여부 구하기
        /// </summary>
        /// <param name="request">요청</param>
        /// <returns>파일 요청 여부</returns>
        private bool IsFileRequest(HttpListenerRequest request)
        {
            string fileExtension = Path.GetExtension(request.Url.AbsolutePath);

            bool isFileRequest = _mimeTypeDictionary.ContainsKey(fileExtension);

            return isFileRequest;
        }

        #endregion
        #region 액션 요청 여부 구하기 - IsActionRequest(request)

        /// <summary>
        /// 액션 요청 여부 구하기
        /// </summary>
        /// <param name="request">요청</param>
        /// <returns>액션 요청 여부</returns>
        private bool IsActionRequest(HttpListenerRequest request)
        {
            bool isActionRequest = request.Url.AbsolutePath.EndsWith(".action");

            return isActionRequest;
        }

        #endregion

        #region 청취하기 - Listen()

        /// <summary>
        /// 청취하기
        /// </summary>
        private void Listen()
        {
            while(this.isRunning)
            {
                HttpListenerContext context = listener.GetContext();

                Thread responseThread = new Thread(new ParameterizedThreadStart(Response));

                responseThread.IsBackground = true;

                this.responseThreadList.Add(responseThread);

                responseThread.Start(new ThreadParameter { Context = context, Thread = responseThread });
            }
        }

        #endregion
        #region 응답하기 - Response(parameter)

        /// <summary>
        /// 응답하기
        /// </summary>
        /// <param name="parameter">매개 변수</param>
        private void Response(object parameter)
        {
            ThreadParameter threadParameter = parameter as ThreadParameter;

            HttpListenerContext  context  = threadParameter.Context;
            HttpListenerRequest  request  = context.Request;
            HttpListenerResponse response = context.Response;

            try
            {
                if(IsDefaultRequest(request))
                {
                    string defaultFilePath = Path.Combine(this.rootPath, "index.html");

                    WriteFile(response, defaultFilePath);
                }
                else if(IsFileRequest(request))
                {
                    string physicalPath = GetPhysicalPath(request.Url.AbsolutePath);

                    string filePath = Path.Combine(this.rootPath, physicalPath);

                    if(File.Exists(filePath))
                    {
                        WriteFile(response, filePath);
                    }
                    else
                    {
                        Write(response, string.Format("요청하신 파일이 존재하지 않습니다 : {0}", request.Url.ToString()), true);
                    }
                }
                else if(IsActionRequest(request))
                {
                    if(ActionRequested != null)
                    {
                        ActionRequested(this, new ActionRequestedEventArgs(this, context));
                    }
                    else
                    {
                        WriteDefaultAction(context);
                    }
                }
                else
                {
                    Write(response, string.Format("요청하신 URL이 존재하지 않습니다 : {0}", request.Url.ToString()), true);
                }
            }
            catch(Exception e)
            {
                Write(response, e.Message + "//" + e.StackTrace, true);
            }
            finally
            {
                this.responseThreadList.Remove(threadParameter.Thread);

                response.Close();
            }
        }

        #endregion
    }
}