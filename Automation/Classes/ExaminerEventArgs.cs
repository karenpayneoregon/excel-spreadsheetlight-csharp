using System;

namespace Automation.Classes
{
    public class ExaminerEventArgs : EventArgs
    {
        public ExaminerEventArgs(string message)
        {
            StatusMessage = message;
        }

        public string StatusMessage { get; set; }
    }
}