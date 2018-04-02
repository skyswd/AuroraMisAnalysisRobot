using System;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Bot.Connector;
using Newtonsoft.Json;
using System.IO;
using System.Web;
using System.Configuration;
using System.Collections.Generic;
using HtmlAgilityPack;

namespace AuroraMisAnalysisRobot
{
    [BotAuthentication]
    public class MessagesController : ApiController
    {
        /// <summary>
        /// POST: api/Messages
        /// Receive a message from a user and reply to it
        /// 實際bot 回覆訊息主程式
        /// </summary>
        public async Task<HttpResponseMessage> Post([FromBody]Activity activity)
        {
            string LUISMessage = activity.Text;//要傳給LUIS的字串

            if (activity.Type == ActivityTypes.Message)
            {
                #region LUIS 設定
                //抓取LUIS 的方法
                //設定LUIS 的KEY
                string strLuisKey = ConfigurationManager.AppSettings["LUISAPIKey"].ToString();
                //設定LUIS 的AppId
                string strLuisAppId = ConfigurationManager.AppSettings["LUISAppId"].ToString();
                //設定本次要傳入的訊息
                string strMessage = HttpUtility.UrlEncode(LUISMessage);
                //Luis連線位址設定
                string strLuisUrl = $"https://api.projectoxford.ai/luis/v1/application?id={strLuisAppId}&subscription-key={strLuisKey}&q={strMessage}";

                ConnectorClient connector = new ConnectorClient(new Uri(activity.ServiceUrl));

                //// 收到文字訊息後，往LUIS送
                WebRequest request = WebRequest.Create(strLuisUrl);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                Stream dataStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(dataStream);
                string json = reader.ReadToEnd();
                Models.LUISResult objLUISRes = JsonConvert.DeserializeObject<Models.LUISResult>(json);
                string strReply = "無法識別的內容";
                #endregion

                if (objLUISRes.intents.Count > 0)
                {
                    string strIntent = objLUISRes.intents[0].intent;

                        #region LUIS 分類
                        switch (strIntent)
                        {
                            case "None":
                                strReply = $"大家好:我是Mis測試機1號，目前沒有此問題的相關訊息，在此提供網上的搜詢訊息提供給看倌參考:\t\n HTTPS://WWW.GOOGLE.COM.TW/search?q=" + strMessage + "&oq=" + strMessage;
                                break;
                            case "黃領巾":
                                strReply = ClawInfoFromWeb();
                                break;
                            default:
                                strReply = @"您在說什麼，我聽不懂~~~(/.\)";
                                break;
                        }
                        #endregion 
                }
                else
                {
                    strReply = "您在說什麼，我聽不懂~~~(@.@)";
                }

                // return our reply to the user
                Activity reply = activity.CreateReply(strReply);
                await connector.Conversations.ReplyToActivityAsync(reply);
            }
            else
            {
                HandleSystemMessage(activity);
            }
            var responses = Request.CreateResponse(HttpStatusCode.OK);
            return responses;
        }

        private Activity HandleSystemMessage(Activity message)
        {
            if (message.Type == ActivityTypes.DeleteUserData)
            {
                // Implement user deletion here
                // If we handle user deletion, return a real message
            }
            else if (message.Type == ActivityTypes.ConversationUpdate)
            {
                // Handle conversation state changes, like members being added and removed
                // Use Activity.MembersAdded and Activity.MembersRemoved and Activity.Action for info
                // Not available in all channels
            }
            else if (message.Type == ActivityTypes.ContactRelationUpdate)
            {
                // Handle add/remove from contact lists
                // Activity.From + Activity.Action represent what happened
            }
            else if (message.Type == ActivityTypes.Typing)
            {
                // Handle knowing tha the user is typing
            }
            else if (message.Type == ActivityTypes.Ping)
            {
            }

            return null;
        }

        #region 黃巾相關 爬蟲
        private string ClawInfoFromWeb()
        {
            var result = "";
            // 1.抓當月清單
            var fightSpace = ClawItccWebSite(DateTime.Now.ToString("yyyy"), DateTime.Now.AddMonths(0).ToString("MM"));

            //2017/12/11 加上跨月時的判斷 
            //==================================================
            //取出下月日期
            var nextMonthDate = System.DateTime.Now.AddMonths(1).ToString("yyyy/MM") + "/01";

            var nextyear = DateTime.Parse(nextMonthDate).ToString("yyyy");
            var nextday = DateTime.Parse(nextMonthDate).ToString("MM");
            //==================================================

            // 2.抓次月清單 & 併入
            fightSpace.AddRange(ClawItccWebSite(nextyear, nextday));

            // 3.篩選搶車位的黃領帶
            var MotelCycle = new MotelCycle();
            var MotelCyclea = fightSpace
                .Where(x => x.topic.Contains("永慶")).ToList();

            result = string.Join("", MotelCyclea.Select(x => x.year + " / " + x.month + " / " + x.day + "\t" + x.topic + "\t" + x.room + "\t\n").ToArray());
            return result;
        }

        private List<TheMeeting> ClawItccWebSite(string searchYear, string searchMonth)
        {
            var result = new List<TheMeeting>();

            // 1.使用 HtmlAgilityPack 分析 XPath
            HtmlWeb webClient = new HtmlWeb();

            // 2.將網址放入在webClient.Load
            HtmlDocument doc = webClient.Load($"http://www.ticc.com.tw/main_ch/EventsCalendar.aspx?uid=146&pid=&YYYY={searchYear}&MM={searchMonth}&DD=01#");

            // 3.取得要分析的 HTML 節點 (div list)
            HtmlNodeCollection divList = doc.DocumentNode.SelectNodes(@"/html/body/div[3]/div/div/div[3]/div");

            foreach (HtmlNode dailyMeetings in divList)
            {
                // 4.div class=list 才有每日的會議清單
                if (!dailyMeetings.GetAttributeValue("class", "").Contains("list"))
                    continue;

                var year = searchYear;
                var month = dailyMeetings.SelectNodes("./div[1]/div[1]")[0].InnerText;
                var day = dailyMeetings.SelectNodes("./div[1]/div[2]")[0].InnerText;

                var customList = dailyMeetings.SelectNodes("./div[2]/div");

                foreach (HtmlNode custom in customList)
                {
                    var topic = custom.SelectNodes("./div[1]")[0].InnerText;
                    var room = custom.SelectNodes("./div[2]")[0].InnerText;

                    result.Add(new TheMeeting() { year = searchYear, month = month, day = day, topic = topic, room = room });
                }
            }

            return result;
        }

        private class TheMeeting
        {
            public string year { get; set; }
            public string month { get; set; }

            public string day { get; set; }

            public string topic { get; set; }

            public string room { get; set; }
        }

        private class MotelCycle
        {
            public string date { get; set; }

            public string topic { get; set; }

            public string room { get; set; }
        }
        #endregion
    }
}