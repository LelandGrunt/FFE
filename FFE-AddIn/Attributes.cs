using System;
using System.Linq;
using System.Reflection;

namespace FFE
{
    // https://pisquare.osisoft.com/thread/2164
    [AttributeUsage(AttributeTargets.Property)]
    public class RequiredAttribute : Attribute
    {
        public string ErrorMessage { get; set; } = "Attribute is required.";

        public static void Check(object objectToCheck)
        {
            var type = objectToCheck.GetType();

            foreach (var prop in type.GetProperties())
            {
                if (prop.GetCustomAttributes(typeof(RequiredAttribute), true).Any())
                {
                    object value = prop.GetValue(objectToCheck, null);

                    if (value == null)
                    {
                        string errorMessage = prop.GetCustomAttribute<RequiredAttribute>(true).ErrorMessage;
                        throw new FfeException(errorMessage);
                    }
                }
            }
        }
    }
}