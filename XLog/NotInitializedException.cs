using System;

namespace XLog
{
  internal class NotInitializedException : Exception
  {
    public NotInitializedException()
      : base("Not initialized.") {
    }
    public NotInitializedException(string message)
      : base(message) {
    }
    public NotInitializedException(string message, Exception inner)
      : base(message, inner) {
    }
  }
}
