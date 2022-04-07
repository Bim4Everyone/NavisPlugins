using System;
using System.IO;

namespace NavisDisciplineChecker {
    public class SimpleLogger : IDisposable {
        private readonly StreamWriter _stream;

        public SimpleLogger(string fileName) {
            _stream = File.CreateText(fileName);
        }

        public void WriteLine(string message) {
            _stream.WriteLine(message);
        }

        #region IDisposable

        protected virtual void Dispose(bool disposing) {
            if(disposing) {
                _stream.Flush();
                _stream.Dispose();
            }
        }

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        #endregion
    }
}