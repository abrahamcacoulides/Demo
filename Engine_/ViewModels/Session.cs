using System;
using System.Collections.Generic;
using System.Linq;
using Engine.EventArgs;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace Engine.ViewModels
{
    public class Session : BaseNotificationClass
    {
        public event EventHandler<MessageEventArgs> OnMessageRaised;

        public string _billsPath;

        private void RaiseMessage(string message)
        {
            OnMessageRaised?.Invoke(this, new MessageEventArgs(message));
        }

        public void GoButton()
        {
             RaiseMessage("Go clicked!!");
        }
    }
}
