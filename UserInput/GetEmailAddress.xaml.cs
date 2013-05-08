using System;
using System.Windows;

namespace UserInput
{
    public partial class MainWindow
    {
        private string _emailAddress;

        public MainWindow()
        {
            InitializeComponent();
        }

        public event Closed onClosed;

        protected virtual void OnOnClosed(string emailAddress)
        {
            Closed handler = onClosed;
            if (handler != null) handler(this, new ClosedArgs(emailAddress));
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _emailAddress = EmailAddress.Text;
            Close();
            OnOnClosed(_emailAddress);
        }

    }

    public delegate void Closed(object sender, ClosedArgs args);

    public class ClosedArgs : EventArgs
    {
        public ClosedArgs(string emailAddress)
        {
            EmailAddress = emailAddress;
        }

        public string EmailAddress { get; private set; }
    }
}