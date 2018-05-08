using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace OutlookRuleMgr.Utilities
{
    public static class AssemblyExtensions
    {
        public static IEnumerable<T> GetImplementations<T>(this Assembly assembly)
        {
            var type = typeof(T);
            return assembly.GetTypes()
                .Where(t => t.IsClass && type.IsAssignableFrom(t))
                .Select(Activator.CreateInstance)
                .Cast<T>();
        }
    }
}
