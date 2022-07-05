using System;
using System.Net;

namespace TestProject
{
    /// <summary>
    /// 액션 요청시 이벤트 인자
    /// </summary>
    public class ActionRequestedEventArgs : EventArgs
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////// Field
        ////////////////////////////////////////////////////////////////////////////////////////// Private

        #region Field

        /// <summary>
        /// 웹 서버
        /// </summary>
        private WebServer server = null;

        /// <summary>
        /// HTTP 청취자 컨텍스트
        /// </summary>
        private HttpListenerContext context = null;

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Property
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 서버 - Server

        /// <summary>
        /// 서버
        /// </summary>
        public WebServer Server
        {
            get
            {
                return this.server;
            }
        }

        #endregion
        #region 컨텍스트 - Context

        /// <summary>
        /// 컨텍스트
        /// </summary>
        public HttpListenerContext Context
        {
            get
            {
                return this.context;
            }
        }

        #endregion

        //////////////////////////////////////////////////////////////////////////////////////////////////// Constructor
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 생성자 - ActionRequestedEventArgs(server, context)

        /// <summary>
        /// 생성자
        /// </summary>
        /// <param name="server">서버</param>
        /// <param name="context">컨텍스트</param>
        public ActionRequestedEventArgs(WebServer server, HttpListenerContext context)
        {
            this.server  = server;
            this.context = context;
        }

        #endregion
    }
}