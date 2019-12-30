using System;

namespace FFE
{
    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    public class FfeFunctionAttribute : Attribute
    {
        public string Provider { get; set; }
    }
}