using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using AngleSharp;
using System.Timers;
using Quartz;
using Quartz.Impl;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            //啟動定時器
            Show();
        }
        //定時器
        public static void Show()
        {
            //创建调度单元

            Task<IScheduler> tsk = StdSchedulerFactory.GetDefaultScheduler();
            IScheduler scheduler = tsk.Result;
            //2.创建一个具体的作业即job (具体的job需要单独在一个文件中执行)

            IJobDetail job = JobBuilder.Create<SendMessageJob>().WithIdentity("完成").Build();
            //3.创建并配置一个触发器即trigger   1s执行一次

            ITrigger _CronTrigger = TriggerBuilder.Create()
              .WithIdentity("定时確認")
              .WithCronSchedule("0 44 15 * * ?") //秒 分 时 某一天 月 周 年(可选参数) //我設定每天下午三點更新
              .Build()
              as ITrigger;
            //4.将job和trigger加入到作业调度池中
            scheduler.ScheduleJob(job, _CronTrigger);
            //5.开启调度
            scheduler.Start();
            Console.ReadLine();
        }

        public class SendMessageJob : IJob
        {
            public async Task Execute(IJobExecutionContext context)
            {

                    //爬蟲 爬網站資訊
                    HttpClient httpClient = new HttpClient();

                    string url = "https://events.ettoday.net/covid19info/index.php7";
                    var responseMessage = await httpClient.GetAsync(url); //發送請求
                                                                          //檢查回應的伺服器狀態StatusCode是否是200 OK
                    if (responseMessage.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        string ClawerResponse = responseMessage.Content.ReadAsStringAsync().Result;//取得內容

                        //Console.WriteLine(response);

                        // 使用AngleSharp時的前置設定
                        var config = Configuration.Default;
                        var contextB = BrowsingContext.New(config);

                        //將我們用httpclient拿到的資料放入res.Content中())
                        var document = await contextB.OpenAsync(res => res.Content(ClawerResponse));

                        //QuerySelector("head")找出<head></head>元素
                        var head = document.QuerySelector("head");
                        //Console.WriteLine(head.ToHtml());

                        //QuerySelector(".entry-content")找出class="entry-content"的所有元素
                        var contents = document.QuerySelectorAll(".table_1 tbody tr td span");
                        string[] CovidVaule = { };
                        int arrayIndex = 0;

                        foreach (var c in contents)
                        {
                            //取得每個元素的TextContent
                            Array.Resize<string>(ref CovidVaule, arrayIndex + 1);
                            CovidVaule[arrayIndex] = Convert.ToString(c.TextContent);

                            arrayIndex++;
                        }

                        ParasGetCovidDetail par = new ParasGetCovidDetail();

                        string today = DateTime.Now.ToString("M/dd");

                        par.Date = today;
                        par.Foreign = CovidVaule[0];
                        par.Local = CovidVaule[1];
                        par.Death = CovidVaule[2];
                        par.TotalForeign = CovidVaule[3];
                        par.TotalLocal = CovidVaule[4];
                        par.TotalDeath = CovidVaule[5];
                        par.Total = CovidVaule[6];


                        //把爬到的資料丟回google sheet
                        //要爬的url
                        string PostUrl = "https://script.google.com/macros/s/AKfycbyT0tyDHNVAAGBBcj97bfAW8RQLlRPdZ4lyZZiVIxfp-2QsBykn5jh_lgCYIBZ5j2iyiA/exec";

                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(PostUrl);
                        request.Method = "POST";
                        request.ContentType = "application/x-www-form-urlencoded";

                        //必須透過ParseQueryString()來建立NameValueCollection物件，之後.ToString()才能轉換成queryString
                        NameValueCollection postParams = System.Web.HttpUtility.ParseQueryString(string.Empty);

                        postParams.Clear();
                        //傳入的參數名稱和值

                        postParams.Add("Date", par.Date);
                        postParams.Add("Foreign", par.Foreign);
                        postParams.Add("Local", par.Local);
                        postParams.Add("Death", par.Death);
                        postParams.Add("TotalForeign", par.TotalForeign);
                        postParams.Add("TotalLocal", par.TotalLocal);
                        postParams.Add("TotalDeath", par.TotalDeath);
                        postParams.Add("Total", par.Total);


                        //Console.WriteLine(postParams.ToString());// 將取得"version=1.0&action=preserveCodeCheck&pCode=pCode&TxID=guid&appId=appId", key和value會自動UrlEncode
                        //要發送的字串轉為byte[] 
                        byte[] byteArray = Encoding.UTF8.GetBytes(postParams.ToString());
                        using (Stream reqStream = request.GetRequestStream())
                        {
                            reqStream.Write(byteArray, 0, byteArray.Length);
                        }//end using

                        //API回傳的字串
                        string responseStr = "";
                        //發出Request
                        using (WebResponse response = request.GetResponse())
                        {
                            using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                            {
                                responseStr = sr.ReadToEnd();
                            }//end using  
                        }

                        //Console.Write(responseStr);//印出回傳字串
                        //Console.ReadKey();//暫停畫面

                    }

                    //Console.ReadKey();

            }
        }

    }
    class test
    {
        public async Task testFunction()
        {
            //爬蟲 爬網站資訊
            HttpClient httpClient = new HttpClient();

            string url = "https://events.ettoday.net/covid19info/index.php7";
            var responseMessage = await httpClient.GetAsync(url); //發送請求
            //檢查回應的伺服器狀態StatusCode是否是200 OK
            if (responseMessage.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string ClawerResponse = responseMessage.Content.ReadAsStringAsync().Result;//取得內容

                //Console.WriteLine(response);

                // 使用AngleSharp時的前置設定
                var config = Configuration.Default;
                var context = BrowsingContext.New(config);

                //將我們用httpclient拿到的資料放入res.Content中())
                var document = await context.OpenAsync(res => res.Content(ClawerResponse));

                //QuerySelector("head")找出<head></head>元素
                var head = document.QuerySelector("head");
                //Console.WriteLine(head.ToHtml());

                //QuerySelector(".entry-content")找出class="entry-content"的所有元素
                var contents = document.QuerySelectorAll(".table_1 tbody tr td span");
                string[] CovidVaule = { };
                int arrayIndex = 0;

                foreach (var c in contents)
                {
                    //取得每個元素的TextContent
                    Array.Resize<string>(ref CovidVaule, arrayIndex + 1);
                    CovidVaule[arrayIndex] = Convert.ToString(c.TextContent);

                    arrayIndex++;
                }

                ParasGetCovidDetail par = new ParasGetCovidDetail();

                string today = DateTime.Now.ToString("M/dd");

                par.Date = today;
                par.Foreign = CovidVaule[0];
                par.Local = CovidVaule[1];
                par.Death = CovidVaule[2];
                par.TotalForeign = CovidVaule[3];
                par.TotalLocal = CovidVaule[4];
                par.TotalDeath = CovidVaule[5];
                par.Total = CovidVaule[6];


                //把爬到的資料丟回google sheet
                //要爬的url
                string PostUrl = "https://script.google.com/macros/s/AKfycbyT0tyDHNVAAGBBcj97bfAW8RQLlRPdZ4lyZZiVIxfp-2QsBykn5jh_lgCYIBZ5j2iyiA/exec";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(PostUrl);
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";

                //必須透過ParseQueryString()來建立NameValueCollection物件，之後.ToString()才能轉換成queryString
                NameValueCollection postParams = System.Web.HttpUtility.ParseQueryString(string.Empty);

                postParams.Clear();
                //傳入的參數名稱和值

                postParams.Add("Date", par.Date);
                postParams.Add("Foreign", par.Foreign);
                postParams.Add("Local", par.Local);
                postParams.Add("Death", par.Death);
                postParams.Add("TotalForeign", par.TotalForeign);
                postParams.Add("TotalLocal", par.TotalLocal);
                postParams.Add("TotalDeath", par.TotalDeath);
                postParams.Add("Total", par.Total);


                //Console.WriteLine(postParams.ToString());// 將取得"version=1.0&action=preserveCodeCheck&pCode=pCode&TxID=guid&appId=appId", key和value會自動UrlEncode
                //要發送的字串轉為byte[] 
                byte[] byteArray = Encoding.UTF8.GetBytes(postParams.ToString());
                using (Stream reqStream = request.GetRequestStream())
                {
                    reqStream.Write(byteArray, 0, byteArray.Length);
                }//end using

                //API回傳的字串
                string responseStr = "";
                //發出Request
                using (WebResponse response = request.GetResponse())
                {
                    using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                    {
                        responseStr = sr.ReadToEnd();
                    }//end using  
                }

                //Console.Write(responseStr);//印出回傳字串
                //Console.ReadKey();//暫停畫面

            }

            //Console.ReadKey();
        }
    }

    public class ParasGetCovidDetail
    {
        public string Date { get; set; }
        public string Foreign { get; set; }
        public string Local { get; set; }
        public string Death { get; set; }
        public string TotalForeign { get; set; }
        public string TotalLocal { get; set; }
        public string TotalDeath { get; set; }
        public string Total { get; set; }


    }
    public class EndResponse
    {
        public bool Success { get; set; }
        public string Result { get; set; }
        public string Message { get; set; }
        public string Data { get; set; }
    }
}
