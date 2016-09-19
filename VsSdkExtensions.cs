//------------------------------------------------------------------------------
// <copyright file="VsSdkExtensions.cs" company="Ben Corp.">
//     Copyright (c) Ben Corp.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace Ben.VisualStudio
{
    /// <summary>
    /// A set of extension methods for various classes in the Visual Studio SDK
    /// </summary>
    public static class VsSdkExtensions
    {
        public static WindowFrameEnumerator GetEnumerator(this IEnumWindowFrames self)
        {
            return new WindowFrameEnumerator(self);
        }

        public static TProperty GetProperty<TProperty>(this IVsWindowFrame windowFrame, __VSFPROPID propertyId)
        {
            object value;
            var result = windowFrame.GetProperty((int)propertyId, out value);

            if (result != VSConstants.S_OK)
            {
                throw new Exception("GetProperty call failed with {result}.");
            }

            return (TProperty)value;
        }

        public static WindowFrameEnumerable GetDocumentWindows(this IVsUIShell shell)
        {
            return new WindowFrameEnumerable(shell);
        }
    
        public static TIService GetService<TService, TIService>(this System.IServiceProvider serviceProvider)
        {
            return (TIService)serviceProvider.GetService(typeof(TService));
        }
    }
}
