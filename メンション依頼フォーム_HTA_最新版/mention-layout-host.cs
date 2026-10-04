// v28.31.2 - Windows-only geometry adapter. No request data or CSV access.
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class MentionLayoutHost {
    [StructLayout(LayoutKind.Sequential)]
    public struct Rect {
        public int Left, Top, Right, Bottom;
        public int Width { get { return Right-Left; } }
        public int Height { get { return Bottom-Top; } }
        public Rect(int x, int y, int w, int h) { Left=x; Top=y; Right=x+w; Bottom=y+h; }
    }
    [StructLayout(LayoutKind.Sequential)]
    struct MonitorInfo { public int Size; public Rect Monitor, Work; public uint Flags; }
    delegate bool EnumProc(IntPtr hwnd, IntPtr state);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr state);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool IsZoomed(IntPtr hwnd);
    [DllImport("user32.dll")] static extern short GetAsyncKeyState(int key);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint flags);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern bool GetMonitorInfo(IntPtr monitor, ref MonitorInfo info);
    [DllImport("user32.dll", SetLastError=true)] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll", SetLastError=true)] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);

    static bool Matches(IntPtr hwnd, int processId, string title) {
        if (!IsWindow(hwnd)) return false;
        uint pid; GetWindowThreadProcessId(hwnd, out pid);
        if (pid != processId) return false;
        StringBuilder text=new StringBuilder(512);
        GetWindowText(hwnd,text,text.Capacity);
        return String.Equals(text.ToString(),title,StringComparison.Ordinal);
    }
    static IntPtr Find(int processId, string title) {
        IntPtr result=IntPtr.Zero; int count=0;
        EnumWindows(delegate(IntPtr hwnd, IntPtr state) {
            if (Matches(hwnd,processId,title)) { result=hwnd; count++; }
            return true;
        },IntPtr.Zero);
        return count==1 ? result : IntPtr.Zero;
    }
    static int Clamp(int value, int low, int high) { return Math.Max(low,Math.Min(value,high)); }

    // Pure geometry, also exercised by the C# regression test on Linux.
    public static Rect Fit(Rect old, Rect work, int width, int height, bool initial) {
        if (work.Width<=16 || work.Height<=16) throw new ArgumentException("Invalid work area");
        int margin=8;
        width=Clamp(width,1,work.Width-margin*2);
        height=Clamp(height,1,work.Height-margin*2);
        int x=Clamp(old.Left,work.Left+margin,work.Right-width-margin);
        int y=Clamp(old.Top,work.Top+margin,work.Bottom-height-margin);
        if (initial) {
            x=work.Left+(work.Width-width)/2;
            y=work.Top+(work.Height-height)/2;
        } else if (height>old.Height) {
            // Restore the earlier upper-center behavior only when expanding.
            int upper=work.Top+margin+(int)Math.Round((work.Height-height-margin*2)*0.38);
            y=Math.Min(y,upper);
        }
        return new Rect(x,y,width,height);
    }

    public static Rect Desired(Rect old, Rect client, Rect work, int[] r) {
        // Native calls run in physical pixels. Derive CSS-to-pixel scale from the
        // same live window, including mixed display scaling and window borders.
        double sx=r[5]>0 ? (double)old.Width/r[5] : (double)client.Width/r[3];
        double sy=r[6]>0 ? (double)old.Height/r[6] : (double)client.Height/r[4];
        if (sx<0.25 || sx>8 || sy<0.25 || sy>8) throw new InvalidOperationException("Invalid display scale");
        int width=(int)Math.Ceiling(r[1]*sx)+Math.Max(0,old.Width-client.Width);
        int height=(int)Math.Ceiling(r[2]*sy)+Math.Max(0,old.Height-client.Height);
        return Fit(old,work,width,height,r[7]==1);
    }

    static bool ReadRequest(string file, out int[] request) {
        request=null;
        try {
            string[] fields=File.ReadAllText(file).Trim().Split('|');
            if (fields.Length!=9) return false;
            int[] values=new int[9];
            for (int i=0;i<9;i++) if (!Int32.TryParse(fields[i],NumberStyles.None,CultureInfo.InvariantCulture,out values[i])) return false;
            if (values[0]<1 || values[0]!=values[8]) return false;
            for (int i=1;i<=4;i++) if (values[i]<1 || values[i]>100000) return false;
            if (values[5]>100000 || values[6]>100000 || values[7]>1) return false;
            request=values; return true;
        } catch (IOException) { return false; }
    }

    public static void Run(string folder, int processId, string title) {
        // Only the specific mshta parent in this logon session is eligible.
        using (Process owner=Process.GetProcessById(processId)) {
            if (!String.Equals(owner.ProcessName,"mshta",StringComparison.OrdinalIgnoreCase) ||
                owner.SessionId!=Process.GetCurrentProcess().SessionId) throw new InvalidOperationException("HTA parent not found");
            IntPtr hwnd=Find(processId,title);
            if (hwnd==IntPtr.Zero) throw new InvalidOperationException("Unique form window not found");
            IntPtr previous=SetThreadDpiAwarenessContext(new IntPtr(-4));
            if (previous==IntPtr.Zero) throw new InvalidOperationException("Physical display coordinates unavailable");
            try {
                File.WriteAllText(Path.Combine(folder,"ready.txt"),"OK");
                int last=0;
                while (!owner.HasExited && Matches(hwnd,processId,title) && !File.Exists(Path.Combine(folder,"stop.txt"))) {
                    if (DateTime.UtcNow-File.GetLastWriteTimeUtc(Path.Combine(folder,"heartbeat.txt"))>TimeSpan.FromSeconds(30)) break;
                    int[] request;
                    if (!ReadRequest(Path.Combine(folder,"request.txt"),out request) || request[0]<=last ||
                        (GetAsyncKeyState(1)&0x8000)!=0) { Thread.Sleep(60); continue; }
                    string status="OK";
                    if (IsIconic(hwnd) || IsZoomed(hwnd)) { status="SKIP"; }
                    else {
                        // Select the monitor BEFORE growth, using this window's largest overlap.
                        MonitorInfo info=new MonitorInfo(); info.Size=Marshal.SizeOf(typeof(MonitorInfo));
                        IntPtr monitor=MonitorFromWindow(hwnd,2);
                        Rect old,client;
                        if (!GetMonitorInfo(monitor,ref info) ||
                            !GetWindowRect(hwnd,out old) || !GetClientRect(hwnd,out client)) throw new InvalidOperationException("Window geometry unavailable");
                        Rect next=Desired(old,client,info.Work,request);
                        int[] current;
                        if (!ReadRequest(Path.Combine(folder,"request.txt"),out current) || current[0]!=request[0]) continue;
                        if (File.Exists(Path.Combine(folder,"stop.txt")) || !Matches(hwnd,processId,title)) break;
                        Rect check;
                        if (IsIconic(hwnd) || IsZoomed(hwnd) || !GetWindowRect(hwnd,out check) ||
                            check.Left!=old.Left || check.Top!=old.Top || check.Width!=old.Width || check.Height!=old.Height ||
                            MonitorFromWindow(hwnd,2)!=monitor) continue;
                        // Size and position change atomically; no activation, z-order or topmost change.
                        if (!SetWindowPos(hwnd,IntPtr.Zero,next.Left,next.Top,next.Width,next.Height,0x0014))
                            throw new InvalidOperationException("Window fit failed: "+Marshal.GetLastWin32Error());
                    }
                    last=request[0];
                    File.WriteAllText(Path.Combine(folder,"result.txt"),last.ToString(CultureInfo.InvariantCulture)+"|"+status);
                }
            } finally { SetThreadDpiAwarenessContext(previous); }
        }
    }
}
