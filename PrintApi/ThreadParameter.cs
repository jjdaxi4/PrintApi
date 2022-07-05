using System.Net;
using System.Threading;

namespace TestProject
{
    /// <summary>
    /// 스레드 매개 변수
    /// </summary>
    public class ThreadParameter
    {
        //////////////////////////////////////////////////////////////////////////////////////////////////// Property
        ////////////////////////////////////////////////////////////////////////////////////////// Public

        #region 컨텍스트 - Context

        /// <summary>
        /// 컨텍스트
        /// </summary>
        public HttpListenerContext Context { get; set; }

        #endregion
        #region 스레드 - Thread

        /// <summary>
        /// 스레드
        /// </summary>
        public Thread Thread { get; set; }

        #endregion
    }
}