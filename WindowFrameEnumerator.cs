//------------------------------------------------------------------------------
// <copyright file="WindowFrameEnumerator.cs" company="Ben Corp.">
//     Copyright (c) Ben Corp.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System.Collections;
using System.Collections.Generic;
using Microsoft.VisualStudio.Shell.Interop;

namespace Ben.VisualStudio
{
    /// <summary>
    /// Provides an <see cref="IEnumerator{IVsWindowFrame}"/> to allow simple
    /// iteration over an <see cref="IEnumWindowFrames"/> object returned by
    /// <see cref="IVsUIShell.GetDocumentWindowEnum(out IEnumWindowFrames)"/>.
    /// </summary>
    public class WindowFrameEnumerator : IEnumerator<IVsWindowFrame>
    {
        public const int BufferSize = 5;

        private IEnumWindowFrames enumWindowFrames;
        private IVsWindowFrame[] buffer;
        private int windowCount = 0;
        private int position = 0;

        public WindowFrameEnumerator(IEnumWindowFrames enumWindowFrames)
        {
            this.enumWindowFrames = enumWindowFrames;
            this.buffer = new IVsWindowFrame[BufferSize];
        }

        public IVsWindowFrame Current
        {
            get
            {
                return this.buffer[position];
            }
        }

        object IEnumerator.Current
        {
            get
            {
                return this.Current;
            }
        }

        public bool MoveNext()
        {
            position++;
            if (position >= windowCount)
            {
                uint framesRetrieved;
                this.enumWindowFrames.Next(BufferSize, this.buffer, out framesRetrieved);
                this.windowCount = (int)framesRetrieved;
                position = 0;

                if (this.windowCount == 0)
                {
                    return false;
                }
            }

            return true;
        }

        public void Reset()
        {
            this.position = 0;
            this.windowCount = 0;
            this.enumWindowFrames.Reset();
        }

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {

                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}
