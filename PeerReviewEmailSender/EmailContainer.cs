using System;
using System.Collections.Generic;
using System.Text;

namespace PeerReviewEmailSender
{
    public class EmailContainer
    {
        private Dictionary<string, List<string>> feedbacksForPresenter =
            new Dictionary<string, List<string>>();

        public void Add(string presenterEmail, string feedback)
        {
            if (!feedbacksForPresenter.ContainsKey(presenterEmail))
            {
                feedbacksForPresenter.Add(presenterEmail, new List<string>());
            }
            feedbacksForPresenter[presenterEmail].Add(feedback);
        }

        public IEnumerable<string> GetPresenterEmails()
        {
            foreach(var e in feedbacksForPresenter.Keys)
                yield return e;
        }

        string messagePrefix = $"Kedves Kolléga!<BR/><BR/>{Environment.NewLine}A beszámolódra az alábbi visszejelzések érkeztek:<BR/><BR/>{Environment.NewLine}{Environment.NewLine}";

        const string feedbackSeparator = "<BR/><BR/>\n\n----------<BR/><BR/>\n\n";

        public string GetMessageBody(string presenterEmail)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(messagePrefix);
            foreach(var f in feedbacksForPresenter[presenterEmail])
            {
                sb.Append(f);
                sb.Append(feedbackSeparator);
            }
            return sb.ToString();
        }
    }
}
