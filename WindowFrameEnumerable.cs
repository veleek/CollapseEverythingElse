//------------------------------------------------------------------------------
// <copyright file="WindowFrameEnumerable.cs" company="Ben Corp.">
//     Copyright (c) Ben Corp.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System.Collections;
using System.Collections.Generic;
using Microsoft.VisualStudio.Shell.Interop;

namespace Ben.VisualStudio
{
    /// <summary>
    /// A class that wraps an <see cref="IEnumWindowFrames"/> object and implements 
    /// <see cref="IEnumerable{IVsWindowFrame}"/> to allow it to be used in a foreach
    /// loop, or other standard situation.  Uses <see cref="WindowFrameEnumerator"/>
    /// to handle actual enumeration.
    /// </summary>
    public class WindowFrameEnumerable : IEnumerable<IVsWindowFrame>
    {
        private IEnumWindowFrames enumWindowFrames;

        public WindowFrameEnumerable(IVsUIShell shell)
        {
            shell.GetDocumentWindowEnum(out this.enumWindowFrames);
        }

        public WindowFrameEnumerable(IEnumWindowFrames enumWindowFrames)
        {
            this.enumWindowFrames = enumWindowFrames;
        }

        public IEnumerator<IVsWindowFrame> GetEnumerator()
        {
            return new WindowFrameEnumerator(this.enumWindowFrames);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }
    }
}
