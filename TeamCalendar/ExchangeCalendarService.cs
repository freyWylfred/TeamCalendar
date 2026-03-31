using System.Net;
using System.Text;
using System.Xml.Linq;

namespace TeamCalendar
{
    /// <summary>
    /// Exchange Server の認証情報
    /// </summary>
    public class ExchangeCredential
    {
        public string ServerUrl { get; set; } = "";
        public string Email { get; set; } = "";
        public string Password { get; set; } = "";
        public string Domain { get; set; } = "";
        public bool IgnoreSslErrors { get; set; }

        public bool IsConfigured =>
            !string.IsNullOrWhiteSpace(ServerUrl) &&
            !string.IsNullOrWhiteSpace(Email) &&
            !string.IsNullOrWhiteSpace(Password);
    }

    /// <summary>
    /// Exchange Web Services (EWS) 経由で予定表を取得するサービス
    /// </summary>
    public sealed class ExchangeCalendarService : IDisposable
    {
        private static readonly XNamespace SoapNs = "http://schemas.xmlsoap.org/soap/envelope/";
        private static readonly XNamespace TypesNs = "http://schemas.microsoft.com/exchange/services/2006/types";
        private static readonly XNamespace MessagesNs = "http://schemas.microsoft.com/exchange/services/2006/messages";

        private readonly HttpClient _httpClient;
        private readonly ExchangeCredential _credential;

        public ExchangeCalendarService(ExchangeCredential credential)
        {
            _credential = credential;

            var handler = new HttpClientHandler
            {
                Credentials = string.IsNullOrWhiteSpace(credential.Domain)
                    ? new NetworkCredential(credential.Email, credential.Password)
                    : new NetworkCredential(credential.Email, credential.Password, credential.Domain),
            };

            if (credential.IgnoreSslErrors)
            {
                handler.ServerCertificateCustomValidationCallback =
                    HttpClientHandler.DangerousAcceptAnyServerCertificateValidator;
            }

            _httpClient = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(30) };
            _httpClient.DefaultRequestHeaders.Add("Accept", "text/xml");
        }

        /// <summary>
        /// 接続テスト（自分の予定表に対して 1 日分の FindItem を実行）
        /// </summary>
        public async Task TestConnectionAsync()
        {
            var today = DateTime.Today;
            await FindAppointmentsAsync(_credential.Email, today, today.AddDays(1));
        }

        /// <summary>
        /// 指定ユーザーの予定を EWS CalendarView で取得する
        /// </summary>
        public async Task<List<AppointmentInfo>> FindAppointmentsAsync(
            string targetEmail, DateTime startDate, DateTime endDate)
        {
            var soap = BuildFindItemSoap(targetEmail, startDate, endDate);

            using var content = new StringContent(soap, Encoding.UTF8, "text/xml");
            using var response = await _httpClient.PostAsync(_credential.ServerUrl, content);

            if (response.StatusCode == HttpStatusCode.Unauthorized)
            {
                throw new UnauthorizedAccessException(
                    "認証に失敗しました (HTTP 401)。メールアドレス・パスワード・ドメインを確認してください。");
            }

            if (!response.IsSuccessStatusCode)
            {
                var body = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException(
                    $"Exchange Server: HTTP {(int)response.StatusCode} {response.ReasonPhrase}\n" +
                    Truncate(body, 500));
            }

            var xml = await response.Content.ReadAsStringAsync();
            return ParseResponse(xml, targetEmail);
        }

        private string BuildFindItemSoap(string targetEmail, DateTime start, DateTime end)
        {
            bool isSelf = string.Equals(
                targetEmail, _credential.Email, StringComparison.OrdinalIgnoreCase);

            var folderId = isSelf
                ? new XElement(TypesNs + "DistinguishedFolderId",
                    new XAttribute("Id", "calendar"))
                : new XElement(TypesNs + "DistinguishedFolderId",
                    new XAttribute("Id", "calendar"),
                    new XElement(TypesNs + "Mailbox",
                        new XElement(TypesNs + "EmailAddress", targetEmail)));

            var envelope = new XElement(SoapNs + "Envelope",
                new XAttribute(XNamespace.Xmlns + "soap", SoapNs),
                new XAttribute(XNamespace.Xmlns + "t", TypesNs),
                new XAttribute(XNamespace.Xmlns + "m", MessagesNs),
                new XElement(SoapNs + "Header",
                    new XElement(TypesNs + "RequestServerVersion",
                        new XAttribute("Version", "Exchange2013"))),
                new XElement(SoapNs + "Body",
                    new XElement(MessagesNs + "FindItem",
                        new XAttribute("Traversal", "Shallow"),
                        new XElement(MessagesNs + "ItemShape",
                            new XElement(TypesNs + "BaseShape", "AllProperties")),
                        new XElement(MessagesNs + "CalendarView",
                            new XAttribute("StartDate", start.ToString("yyyy-MM-ddT00:00:00")),
                            new XAttribute("EndDate", end.ToString("yyyy-MM-ddT00:00:00"))),
                        new XElement(MessagesNs + "ParentFolderIds", folderId))));

            return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + envelope.ToString();
        }

        private static List<AppointmentInfo> ParseResponse(string xml, string ownerEmail)
        {
            var doc = XDocument.Parse(xml);
            var appointments = new List<AppointmentInfo>();

            var responseCode = doc.Descendants(MessagesNs + "ResponseCode").FirstOrDefault();
            if (responseCode is not null && responseCode.Value != "NoError")
            {
                var msg = doc.Descendants(MessagesNs + "MessageText").FirstOrDefault()?.Value ?? "";
                throw new InvalidOperationException(
                    $"EWS エラー: {responseCode.Value}\n{msg}");
            }

            foreach (var item in doc.Descendants(TypesNs + "CalendarItem"))
            {
                string subject = item.Element(TypesNs + "Subject")?.Value ?? "(件名なし)";
                DateTime.TryParse(item.Element(TypesNs + "Start")?.Value, out var startDt);
                DateTime.TryParse(item.Element(TypesNs + "End")?.Value, out var endDt);
                string organizer = item.Element(TypesNs + "Organizer")
                    ?.Element(TypesNs + "Mailbox")
                    ?.Element(TypesNs + "Name")?.Value ?? "(取得失敗)";
                string location = item.Element(TypesNs + "Location")?.Value ?? "";
                string responseType = item.Element(TypesNs + "MyResponseType")?.Value ?? "Unknown";

                int duration = (int)(endDt - startDt).TotalMinutes;
                int responseStatus = MapResponseType(responseType);

                appointments.Add(new AppointmentInfo
                {
                    Owner = ownerEmail,
                    Subject = subject,
                    Start = startDt,
                    End = endDt,
                    Duration = duration,
                    Organizer = organizer,
                    Location = location,
                    Status = GetStatusText(responseStatus),
                    ResponseStatus = responseStatus,
                });
            }

            return appointments;
        }

        private static int MapResponseType(string type) => type switch
        {
            "Organizer" => 1,
            "Tentative" or "TentativelyAccepted" => 2,
            "Accept" or "Accepted" => 3,
            "Decline" or "Declined" => 4,
            "NoResponseReceived" => 5,
            _ => 0,
        };

        private static string GetStatusText(int status) => status switch
        {
            0 => "未設定",
            1 => "主催者",
            2 => "任意",
            3 => "承認",
            4 => "辞退",
            5 => "未応答",
            _ => $"不明({status})",
        };

        private static string Truncate(string text, int max) =>
            text.Length > max ? text[..max] + "..." : text;

        public void Dispose() => _httpClient.Dispose();
    }
}
