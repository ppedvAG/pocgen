using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppedv.pocgen.Domain.Models
{
    public static class MessagingCenter // ehe "PseudoMessagingCenter" ;)
    {
        static MessagingCenter()
        {
            subscribers = new Dictionary<string, List<Action<object, System.EventArgs>>>();
        }

        private static Dictionary<string, List<Action<object, EventArgs>>> subscribers;
        public static void Subscribe(string messageID, Action<object,EventArgs> action)
        {
            if(subscribers.ContainsKey(messageID))
            {
                subscribers[messageID].Add(action);
            }
            else
            {
                subscribers.Add(messageID, new List<Action<object, EventArgs>> { action });
            }
        }

        public static void Unsubscribe(string messageID, Action<object, EventArgs> action)
        {
            if (subscribers.ContainsKey(messageID))
            {
                subscribers[messageID].Remove(action);
            }
        }

        public static void Send<T>(T sender, string messageID, EventArgs arg)
        {
            if (subscribers.ContainsKey(messageID))
            {
                foreach (Action<object, EventArgs> action in subscribers[messageID])
                    action?.Invoke(sender, arg);
            }
        }
    }
}
