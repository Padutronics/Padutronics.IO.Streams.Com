using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace Padutronics.IO.Streams.Com;

public sealed class ComStreamWrapper : StreamWrapper, IStream
{
    private const int STREAM_SEEK_SET = 0;
    private const int STREAM_SEEK_CUR = 1;
    private const int STREAM_SEEK_END = 2;

    public ComStreamWrapper(Stream wrappee) :
        base(wrappee)
    {
    }

    public void Clone(out IStream ppstm)
    {
        throw new NotImplementedException();
    }

    public void Commit(int grfCommitFlags)
    {
        throw new NotImplementedException();
    }

    private SeekOrigin Convert(int origin)
    {
        return origin switch
        {
            STREAM_SEEK_CUR => SeekOrigin.Current,
            STREAM_SEEK_END => SeekOrigin.End,
            STREAM_SEEK_SET => SeekOrigin.Begin,
            _ => throw new ArgumentOutOfRangeException(nameof(origin), origin, $"Seek origin value {origin} is out of range."),
        };
    }

    public void CopyTo(IStream pstm, long cb, IntPtr pcbRead, IntPtr pcbWritten)
    {
        throw new NotImplementedException();
    }

    public void LockRegion(long libOffset, long cb, int dwLockType)
    {
        throw new NotImplementedException();
    }

    public void Read(byte[] pv, int cb, IntPtr pcbRead)
    {
        int bytesRead = base.Read(pv, 0, cb);

        if (pcbRead != IntPtr.Zero)
        {
            Marshal.WriteInt32(pcbRead, bytesRead);
        }
    }

    public void Revert()
    {
        throw new NotImplementedException();
    }

    public void Seek(long dlibMove, int dwOrigin, IntPtr plibNewPosition)
    {
        int position = (int)Seek(dlibMove, Convert(dwOrigin));

        if (plibNewPosition != IntPtr.Zero)
        {
            Marshal.WriteInt32(plibNewPosition, position);
        }
    }

    public void SetSize(long libNewSize)
    {
        throw new NotImplementedException();
    }

    public void Stat(out STATSTG pstatstg, int grfStatFlag)
    {
        pstatstg = new STATSTG { cbSize = Length };
    }

    public void UnlockRegion(long libOffset, long cb, int dwLockType)
    {
        throw new NotImplementedException();
    }

    public void Write(byte[] pv, int cb, IntPtr pcbWritten)
    {
        Write(pv, 0, cb);

        if (pcbWritten != IntPtr.Zero)
        {
            Marshal.WriteInt32(pcbWritten, cb);
        }
    }
}