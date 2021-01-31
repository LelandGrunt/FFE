using System;
using System.Runtime.Serialization;

namespace FFE
{
    [Serializable]
    public class FfeException : Exception
    {
        public FfeException() { }

        public FfeException(string message) : base(message) { }

        public FfeException(string message, Exception innerException) : base(message, innerException) { }

        protected FfeException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info == null)
                throw new ArgumentNullException("info");
            base.GetObjectData(info, context);
        }
    }

    [Serializable]
    public class WebException : Exception
    {
        public WebException() { }

        public WebException(string message) : base(message) { }

        public WebException(string message, Exception innerException) : base(message, innerException) { }

        protected WebException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info == null)
                throw new ArgumentNullException("info");
            base.GetObjectData(info, context);
        }
    }

    [Serializable]
    public class XPathException : Exception
    {
        public string Html { get; set; }

        public string XPath { get; set; }

        public XPathException() { }

        public XPathException(string message) : base(message) { }

        public XPathException(string message, Exception innerException) : base(message, innerException) { }

        protected XPathException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            Html = info.GetString("Html");
            XPath = info.GetString("XPath");
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info == null)
                throw new ArgumentNullException("info");
            info.AddValue("Html", Html);
            info.AddValue("XPath", XPath);
            base.GetObjectData(info, context);
        }
    }

    [Serializable]
    public class CssSelectorException : Exception
    {
        public string Html { get; set; }

        public string CssSelector { get; set; }

        public CssSelectorException() { }

        public CssSelectorException(string message) : base(message) { }

        public CssSelectorException(string message, Exception innerException) : base(message, innerException) { }

        protected CssSelectorException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            Html = info.GetString("Html");
            CssSelector = info.GetString("CssSelector");
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info == null)
                throw new ArgumentNullException("info");
            info.AddValue("Html", Html);
            info.AddValue("CssSelector", CssSelector);
            base.GetObjectData(info, context);
        }
    }

    [Serializable]
    public class JsonPathException : Exception
    {
        public string Json { get; set; }

        public string JsonPath { get; set; }

        public JsonPathException() { }

        public JsonPathException(string message) : base(message) { }

        public JsonPathException(string message, Exception innerException) : base(message, innerException) { }

        protected JsonPathException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            Json = info.GetString("Json");
            JsonPath = info.GetString("JsonPath");
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info == null)
                throw new ArgumentNullException("info");
            info.AddValue("Json", Json);
            info.AddValue("JsonPath", JsonPath);
            base.GetObjectData(info, context);
        }
    }
}