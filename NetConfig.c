/* ============================================================================
 * NetConfig.c - RTL8152B Network Config & Diagnostic Tool
 * WinCE 7.0 ARMv4I, C89
 *
 * Author      : SweatierUnicorn
 * GitHub      : https://github.com/SweatierUnicorn
 * Distribution: Provided free of charge by the author.
 *
 * Buttons: DHCP | Static | Status | Ping | Log | Clr | Save | Echo
 *
 * DHCP:   Setup DHCP mode, rebind adapter, wait, show status
 * Static: Open touch-friendly dialog to edit IP/Mask/GW/DNS, apply, rebind
 * Status: Show all adapters (IP, MAC, gateway, DHCP mode)
 * Ping:   Ping gateway + 8.8.8.8
 * Log:    Show latest rtl8152_N.log from \MediaCard
 * Clr:    Clear listbox
 * Save:   Save listbox to \MediaCard\netconfig_save.txt
 * ============================================================================ */

#pragma warning(disable: 4996)
#include <winsock2.h>
#include <windows.h>
#include <iphlpapi.h>
#include <ntddndis.h>
#include <winioctl.h>
#include <string.h>

/* VirtualCopy prototype - exported by COREDLL but not in public SDK */
BOOL VirtualCopy(LPVOID lpvDest, LPVOID lpvSrc, DWORD cbSize, DWORD fdwProtect);

#ifndef IOCTL_NDIS_REBIND_ADAPTER
#define IOCTL_NDIS_REBIND_ADAPTER 0x17002E
#endif

#ifndef GWL_HINSTANCE
#define GWL_HINSTANCE (-6)
#endif

#define LOG_FILE L"\\MediaCard\\netconfig_log.txt"

#define IDC_LIST       100
#define IDC_BTN_DHCP   101
#define IDC_BTN_STATIC 102
#define IDC_BTN_STATUS 103
#define IDC_BTN_PING   104
#define IDC_BTN_LOG    105
#define IDC_BTN_CLEAR  106
#define IDC_BTN_SAVE   107
#define IDC_BTN_PINGSRV 108
#define IDC_BTN_PROBE  109
#define IDC_BTN_STAGE  110
#define IDC_BTN_KICK   111
#define IDC_BTN_TESTDLL 112
#define IDC_BTN_VBUS   113
#define IDC_BTN_DNS    114
#define IDC_BTN_HTTP   115
#define IDC_BTN_DL10S  116

#define RTL_WIN_DLL     L"\\Windows\\rtl8152.dll"

#define NETCFG_STARTUP_MODE_NONE   0
#define NETCFG_STARTUP_MODE_STATIC 1
#define NETCFG_STARTUP_MODE_DHCP   2
#define NETCFG_DEFAULT_STARTUP_MODE NETCFG_STARTUP_MODE_STATIC

static HWND g_hList;
static HWND g_hWnd;
static volatile LONG g_bBusy = 0;

static const GUID AX_DEVCLASS_NDIS_GUID =
    {0x98C5250D, 0xC29A, 0x4985, {0xAE, 0x5F, 0xAF, 0xE5, 0x36, 0x7E, 0x50, 0x06}};
static const GUID AX_DEVCLASS_USB_DEVICE_GUID =
    {0x08090549, 0xCE4C, 0x45CE, {0x8A, 0x23, 0xFA, 0x45, 0xF1, 0x16, 0x18, 0x6C}};

typedef HMODULE (WINAPI *PFN_LOAD_DRIVER)(LPCWSTR);

typedef struct {
    GUID  guidDevClass;
    DWORD dwReserved;
    BOOL  fAttached;
    int   cbName;
    TCHAR szName[1];
} AX_DEVDETAIL;

static void AddLog(const WCHAR *msg);
static void RegSetSz(HKEY hKey, const WCHAR *name, const WCHAR *val);
static void RegSetDw(HKEY hKey, const WCHAR *name, DWORD val);
static void RegSetMsz1(HKEY hKey, const WCHAR *name, const WCHAR *val);
static void AddRegSzValue(const WCHAR *keyPath, const WCHAR *valueName);
static void AddRegDwValue(const WCHAR *keyPath, const WCHAR *valueName);
static void DumpSubkeysSimple(const WCHAR *keyPath);
static void DumpKeySnapshot(const WCHAR *keyPath);
static void DumpSubkeysSnapshot(const WCHAR *keyPath, DWORD maxSubkeys);
static void ProbeDeviceOpen(const WCHAR *label, const WCHAR *deviceName);
static BOOL QueryRegSz(HKEY rootKey, const WCHAR *keyPath, const WCHAR *valueName,
                       WCHAR *buf, DWORD cchBuf);
static void ProbeAndroidDriverRoute(void);
static void DumpDummyClassLoadState(void);
static void ArmDummyClassLoadRequests(void);
static void ClearProbeState(void);
static void DumpActiveDriversSummary(void);
static void ProbeUsbNamedEvents(void);
static void ProbePnpNotifications(void);
static void ProbeEhciLiveDuringWait(void);
static BOOL ResolveHostA(const char *host, DWORD *ipNBO);
static BOOL ConnectTcpHost(const char *host, USHORT port, DWORD timeoutMs,
                           SOCKET *outSock, DWORD *outIpNBO);
static BOOL HttpSimpleGet(const char *host, const char *path,
                          DWORD *statusCode, DWORD *elapsedMs, DWORD *bodyBytes);
static BOOL HttpDownloadForMs(const char *host, const char *path, DWORD durationMs,
                              DWORD *elapsedMs, DWORD *bodyBytes, DWORD *statusCode);
static void AtoW(const char *src, WCHAR *dst, int maxLen);
static BOOL QueryRegMultiSzFirst(HKEY rootKey, const WCHAR *keyPath,
                                 const WCHAR *valueName, WCHAR *buf, DWORD cchBuf);
static BOOL ParseIpv4W(const WCHAR *src, int octets[4]);
static void LoadStaticCfgFromRegistry(void);
static BOOL FindRtl8152Adapter(DWORD *ifIdx, WCHAR *ipBuf, DWORD cchIpBuf,
                               BOOL *dhcpEnabled);
static BOOL IsZeroIpStringW(const WCHAR *ip);
static void LogResolvedHost(const char *host);
static DWORD WINAPI AutoStartupThread(LPVOID param);
static DWORD WINAPI NetCheckThread(LPVOID param);
static DWORD WINAPI DnsTestThread(LPVOID param);
static DWORD WINAPI HttpTestThread(LPVOID param);
static DWORD WINAPI Download10sThread(LPVOID param);

static BOOL FileExistsW(const WCHAR *path)
{
    return (GetFileAttributes(path) != 0xFFFFFFFF);
}

static void RegSetDllValue(const WCHAR *keyPath, const WCHAR *dllValue)
{
    HKEY hKey;
    DWORD disp;
    LONG rc;
    WCHAR msg[320];

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, NULL, 0, 0,
                        NULL, &hKey, &disp);
    if (rc != ERROR_SUCCESS) {
        wsprintfW(msg, L"REG create FAIL: HKLM\\%s rc=%d", keyPath, rc);
        AddLog(msg);
        return;
    }

    rc = RegSetValueEx(hKey, L"Dll", 0, REG_SZ,
                       (const BYTE*)dllValue,
                       (DWORD)((wcslen(dllValue) + 1) * sizeof(WCHAR)));
    RegSetValueEx(hKey, L"DLL", 0, REG_SZ,
                  (const BYTE*)dllValue,
                  (DWORD)((wcslen(dllValue) + 1) * sizeof(WCHAR)));
    RegSetValueEx(hKey, L"Prefix", 0, REG_SZ,
                  (const BYTE*)L"RTL",
                  (DWORD)((wcslen(L"RTL") + 1) * sizeof(WCHAR)));
    RegCloseKey(hKey);
    wsprintfW(msg, L"REG set HKLM\\%s\\Dll/DLL/Prefix = %s rc=%d", keyPath, dllValue, rc);
    AddLog(msg);
}

static void ArmAxLoadClientBase(const WCHAR *basePath)
{
    static const WCHAR *suffixes[] = {
        L"\\Default\\Default\\RTL8152B_Class",
        L"\\Default\\255\\RTL8152B_Class",
        L"\\Default\\255_255\\RTL8152B_Class",
        L"\\Default\\255_255_0\\RTL8152B_Class",
        L"\\0_0_0\\255\\RTL8152B_Class",
        L"\\0_0_0\\255_255\\RTL8152B_Class",
        L"\\0_0_0\\255_255_0\\RTL8152B_Class",
        L"\\239_2_1\\255\\RTL8152B_Class",
        L"\\239_2_1\\255_255\\RTL8152B_Class",
        L"\\239_2_1\\255_255_0\\RTL8152B_Class"
    };
    WCHAR path[256];
    int i;

    for (i = 0; i < (int)(sizeof(suffixes) / sizeof(suffixes[0])); ++i) {
        wsprintfW(path, L"%s%s", basePath, suffixes[i]);
        RegSetDllValue(path, L"rtl8152.dll");
    }
}

static void ProbeAxLoadClientBase(const WCHAR *basePath)
{
    static const WCHAR *valueSuffixes[] = {
        L"\\Default\\Default\\RTL8152B_Class",
        L"\\Default\\255\\RTL8152B_Class",
        L"\\Default\\255_255\\RTL8152B_Class",
        L"\\Default\\255_255_0\\RTL8152B_Class",
        L"\\0_0_0\\255\\RTL8152B_Class",
        L"\\0_0_0\\255_255\\RTL8152B_Class",
        L"\\0_0_0\\255_255_0\\RTL8152B_Class",
        L"\\239_2_1\\255\\RTL8152B_Class",
        L"\\239_2_1\\255_255\\RTL8152B_Class",
        L"\\239_2_1\\255_255_0\\RTL8152B_Class"
    };
    static const WCHAR *dumpSuffixes[] = {
        L"\\Default\\Default",
        L"\\Default\\255",
        L"\\Default\\255_255",
        L"\\Default\\255_255_0",
        L"\\0_0_0\\255",
        L"\\0_0_0\\255_255",
        L"\\0_0_0\\255_255_0",
        L"\\239_2_1\\255",
        L"\\239_2_1\\255_255",
        L"\\239_2_1\\255_255_0"
    };
    WCHAR path[256];
    int i;

    for (i = 0; i < (int)(sizeof(valueSuffixes) / sizeof(valueSuffixes[0])); ++i) {
        wsprintfW(path, L"%s%s", basePath, valueSuffixes[i]);
        AddRegSzValue(path, L"Dll");
        AddRegSzValue(path, L"DLL");
    }

    for (i = 0; i < (int)(sizeof(dumpSuffixes) / sizeof(dumpSuffixes[0])); ++i) {
        wsprintfW(path, L"%s%s", basePath, dumpSuffixes[i]);
        DumpSubkeysSimple(path);
    }
}

static void ProbeAndroidDriverRoute(void)
{
    WCHAR routePath[260];
    WCHAR routeName[128];
    WCHAR fullPath[320];
    WCHAR clientPath[256];

    AddLog(L"--- Android USB route ---");
    AddRegSzValue(L"Drivers\\USB\\LoadClients", L"AndroidDriverIdPath");
    AddRegSzValue(L"Drivers\\USB\\LoadClients", L"AndroidDriverIdName");
    AddRegDwValue(L"Drivers\\USB\\AndroidDevType\\0", L"AndroidDevType");
    AddRegDwValue(L"Drivers\\USB\\AndroidDevType\\1", L"AndroidDevType");
    AddRegDwValue(L"Drivers\\USB\\AndroidDevType\\2", L"AndroidDevType");
    AddRegDwValue(L"Drivers\\USB\\AndroidDevType\\3", L"AndroidDevType");

    if (!QueryRegSz(HKEY_LOCAL_MACHINE, L"Drivers\\USB\\LoadClients",
                    L"AndroidDriverIdPath", routePath,
                    sizeof(routePath) / sizeof(routePath[0]))) {
        return;
    }

    DumpSubkeysSimple(routePath);

    if (!QueryRegSz(HKEY_LOCAL_MACHINE, L"Drivers\\USB\\LoadClients",
                    L"AndroidDriverIdName", routeName,
                    sizeof(routeName) / sizeof(routeName[0]))) {
        return;
    }

    wsprintfW(fullPath, L"%s\\%s", routePath, routeName);
    AddRegSzValue(fullPath, L"Dll");
    AddRegSzValue(fullPath, L"DLL");
    AddRegSzValue(fullPath, L"Prefix");

    wsprintfW(clientPath, L"Drivers\\USB\\ClientDrivers\\%s", routeName);
    AddRegSzValue(clientPath, L"Dll");
    AddRegSzValue(clientPath, L"DLL");
    AddRegSzValue(clientPath, L"Prefix");
}

static void AppendTcpipBindValue(void)
{
    HKEY hKey;
    WCHAR existing[512];
    DWORD type;
    DWORD cbData;
    LONG rc;
    WCHAR *p;
    int found;

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\Tcpip\\Linkage", 0, NULL, 0,
                        0, NULL, &hKey, NULL);
    if (rc != ERROR_SUCCESS) {
        AddLog(L"Bind FAIL: cannot open Comm\\Tcpip\\Linkage");
        return;
    }

    memset(existing, 0, sizeof(existing));
    type = 0;
    cbData = sizeof(existing);
    rc = RegQueryValueEx(hKey, L"Bind", NULL, &type, (BYTE*)existing, &cbData);
    if (rc == ERROR_SUCCESS && type == REG_MULTI_SZ && cbData >= 2 * sizeof(WCHAR)) {
        p = existing;
        found = 0;
        while (*p) {
            if (wcscmp(p, L"RTL8152B1") == 0) {
                found = 1;
                break;
            }
            p += wcslen(p) + 1;
        }
        if (!found) {
            wcscpy(p, L"RTL8152B1");
            p += 10;
            *p = L'\0';
            cbData = (DWORD)((BYTE*)(p + 1) - (BYTE*)existing);
            RegSetValueEx(hKey, L"Bind", 0, REG_MULTI_SZ, (const BYTE*)existing, cbData);
            AddLog(L"Bind OK: appended RTL8152B1 to Comm\\Tcpip\\Linkage");
        } else {
            AddLog(L"Bind OK: RTL8152B1 already present in Comm\\Tcpip\\Linkage");
        }
    } else {
        WCHAR msz[16];
        memset(msz, 0, sizeof(msz));
        wcscpy(msz, L"RTL8152B1");
        RegSetValueEx(hKey, L"Bind", 0, REG_MULTI_SZ, (const BYTE*)msz,
                      11 * sizeof(WCHAR));
        AddLog(L"Bind OK: created Comm\\Tcpip\\Linkage with RTL8152B1");
    }

    RegCloseKey(hKey);
}

static void DoKickNdis(void)
{
    HANDLE hNdis;
    WCHAR regMsz[24];
    WCHAR bindMsz[24];
    DWORD cbOut;
    BOOL ok;
    WCHAR msg[192];

    AddLog(L"=== NDIS Kick ===");
    AddLog(L"Kick note: this does NOT trigger USB re-enumeration");

    hNdis = CreateFile(L"NDS0:", GENERIC_READ | GENERIC_WRITE,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_ALWAYS, 0, NULL);
    if (hNdis == INVALID_HANDLE_VALUE) {
        wsprintfW(msg, L"Kick FAIL: NDS0 open err=%d", GetLastError());
        AddLog(msg);
        return;
    }

    memset(regMsz, 0, sizeof(regMsz));
    wcscpy(regMsz, L"RTL8152B");
    wcscpy(regMsz + 9, L"RTL8152B1");
    cbOut = 0;
    ok = DeviceIoControl(hNdis, 0x170026,
                         regMsz, 20 * sizeof(WCHAR), NULL, 0, &cbOut, NULL);
    wsprintfW(msg, L"REGISTER_ADAPTER: %s err=%d", ok ? L"OK" : L"FAIL", GetLastError());
    AddLog(msg);

    memset(bindMsz, 0, sizeof(bindMsz));
    wcscpy(bindMsz, L"RTL8152B1");
    wcscpy(bindMsz + 10, L"TCPIP");
    cbOut = 0;
    ok = DeviceIoControl(hNdis, 0x170032,
                         bindMsz, 17 * sizeof(WCHAR), NULL, 0, &cbOut, NULL);
    wsprintfW(msg, L"BIND_PROTOCOLS: %s err=%d", ok ? L"OK" : L"FAIL", GetLastError());
    AddLog(msg);

    CloseHandle(hNdis);
    AddLog(L"=== NDIS Kick Done ===");
}

/* ============================================================================
 * DoEhciTakeover - Access EHCI registers to clear Port Owner bit (bit 13)
 * so EHCI handles the external port instead of OHCI.
 *
 * Strategy:
 *   A) RegBase from registry = BSP static kernel virtual address. After
 *      SetKMode(TRUE) we should be able to read/write it directly.
 *   B) If A fails, try VirtualCopy of physical HWMemBase/MemBase address.
 *   C) If B fails, try VirtualCopy with different flag combinations.
 *
 * EHCI capability registers start at base. PORTSC = base + CAPLENGTH + 0x44.
 * On the target EHCI layout, CAPLENGTH is 0x20, so PORTSC is at offset 0x64.
 * ============================================================================ */

/* Helper: try reading DWORD from addr, returns 1 if accessible */
static int TryReadDword(volatile DWORD *addr, DWORD *outVal)
{
    /* Use IsBadReadPtr to check before access */
    if (IsBadReadPtr((const void*)addr, 4))
        return 0;
    *outVal = *addr;
    return 1;
}

/* Helper: run EHCI register logic on a mapped base pointer */
static void DoEhciRegisterWork(volatile BYTE *base, DWORD ohciRegBase,
                               int tryOhci)
{
    WCHAR buf[320];
    DWORD capReg, capLen, hciVer, hcsParams, hccParams;
    int nPorts, i, targetPort;
    DWORD opBase, usbCmd, usbSts;

    capReg    = *(volatile DWORD *)(base + 0);
    capLen    = capReg & 0xFF;
    hciVer    = (capReg >> 16) & 0xFFFF;
    hcsParams = *(volatile DWORD *)(base + 4);
    hccParams = *(volatile DWORD *)(base + 8);
    nPorts    = hcsParams & 0xF;
    opBase    = capLen;

    wsprintfW(buf, L"CAPLENGTH=0x%02X HCIVERSION=0x%04X", capLen, hciVer);
    AddLog(buf);
    wsprintfW(buf, L"HCSPARAMS=0x%08X N_PORTS=%d", hcsParams, nPorts);
    AddLog(buf);
    wsprintfW(buf, L"HCCPARAMS=0x%08X", hccParams);
    AddLog(buf);

    if (hciVer != 0x0100 && hciVer != 0x0110 && hciVer != 0x0020) {
        wsprintfW(buf, L"WARNING: HCI ver 0x%04X unexpected", hciVer);
        AddLog(buf);
    }
    if (capLen < 0x08 || capLen > 0x40 || nPorts == 0 || nPorts > 15) {
        wsprintfW(buf, L"FAIL: bad capLen=%u or nPorts=%d, wrong mapping",
                  capLen, nPorts);
        AddLog(buf);
        return;
    }

    usbCmd = *(volatile DWORD *)(base + opBase + 0x00);
    usbSts = *(volatile DWORD *)(base + opBase + 0x04);
    wsprintfW(buf, L"USBCMD=0x%08X USBSTS=0x%08X", usbCmd, usbSts);
    AddLog(buf);

    /* Read PORTSC for each port */
    targetPort = -1;
    for (i = 0; i < nPorts && i < 4; i++) {
        DWORD ps = *(volatile DWORD *)(base + opBase + 0x44 + 4 * i);
        wsprintfW(buf,
            L"PORTSC[%d]=0x%08X CCS=%u PE=%u OCA=%u PP=%u "
            L"PO=%u Rst=%u Susp=%u LS=%u",
            i, ps,
            (ps >> 0) & 1,    /* Current Connect Status */
            (ps >> 2) & 1,    /* Port Enabled */
            (ps >> 4) & 1,    /* Over-current Active */
            (ps >> 12) & 1,   /* Port Power */
            (ps >> 13) & 1,   /* Port Owner */
            (ps >> 8) & 1,    /* Port Reset */
            (ps >> 7) & 1,    /* Suspend */
            (ps >> 10) & 3);  /* Line Status */
        AddLog(buf);
        if ((ps & 0x2000) && targetPort < 0)
            targetPort = i;
    }

    /* Clear Port Owner on first port that has it set */
    if (targetPort >= 0) {
        DWORD portscOff = opBase + 0x44 + 4 * targetPort;
        DWORD ps, newVal;
        int sec;

        ps = *(volatile DWORD *)(base + portscOff);
        wsprintfW(buf,
            L"*** PORTSC[%d]=0x%08X PO=1 (OHCI owns). Clearing bit13...",
            targetPort, ps);
        AddLog(buf);

        /* EHCI PORTSC write: bits 1,3,5 are RWC, write 0 to preserve */
        newVal = ps;
        newVal &= ~0x002Au; /* Clear RWC bits 1,3,5 */
        newVal &= ~0x2000u; /* Clear Port Owner */
        *(volatile DWORD *)(base + portscOff) = newVal;
        Sleep(100);

        ps = *(volatile DWORD *)(base + portscOff);
        wsprintfW(buf,
            L"PORTSC[%d] after = 0x%08X PO=%u CCS=%u PE=%u PP=%u LS=%u",
            targetPort, ps,
            (ps >> 13) & 1, ps & 1, (ps >> 2) & 1,
            (ps >> 12) & 1, (ps >> 10) & 3);
        AddLog(buf);

        if (ps & 0x2000)
            AddLog(L"*** PO reappeared - hw/driver override. Fighting 20s...");
        else
            AddLog(L"*** PO CLEARED! EHCI controls port. High-Speed OK");
        AddLog(L"PLUG ADAPTER NOW...");

        /* Monitor 20s */
        for (sec = 0; sec < 20; sec++) {
            DWORD p = *(volatile DWORD *)(base + portscOff);
            if (p & 0x2000) {
                DWORD nv = p & ~0x202Au; /* clear RWC + PO */
                *(volatile DWORD *)(base + portscOff) = nv;
                p = *(volatile DWORD *)(base + portscOff);
                wsprintfW(buf,
                    L"  [%02ds] PS=0x%08X PO=%u CCS=%u PE=%u (re-cleared)",
                    sec, p, (p>>13)&1, p&1, (p>>2)&1);
            } else {
                wsprintfW(buf,
                    L"  [%02ds] PS=0x%08X PO=%u CCS=%u PE=%u LS=%u",
                    sec, p, (p>>13)&1, p&1, (p>>2)&1, (p>>10)&3);
            }
            AddLog(buf);
            if ((p & 1) && (p & 4) && !(p & 0x2000))
                AddLog(L"  ** DEVICE CONNECTED + ENABLED AT EHCI! **");
            /* Check DummyClassLoad */
            {
                HKEY hDcl; DWORD lr = 0, sz;
                int hi;
                for (hi = 1; hi <= 3; hi++) {
                    WCHAR dp[64];
                    wsprintfW(dp, L"Drivers\\USB\\DummyClassLoad\\%d", hi);
                    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, dp,
                                     0, KEY_READ, &hDcl) == 0) {
                        sz = 4; lr = 0;
                        RegQueryValueEx(hDcl, L"LoadRequest", NULL,
                                        NULL, (BYTE*)&lr, &sz);
                        if (lr) {
                            wsprintfW(buf, L"  ** DCL\\%d: LoadReq=%u",
                                      hi, lr);
                            AddLog(buf);
                        }
                        RegCloseKey(hDcl);
                    }
                }
            }
            Sleep(1000);
        }
        ps = *(volatile DWORD *)(base + portscOff);
        wsprintfW(buf, L"Final PORTSC[%d]=0x%08X PO=%u CCS=%u PE=%u",
                  targetPort, ps, (ps>>13)&1, ps&1, (ps>>2)&1);
        AddLog(buf);
    } else {
        AddLog(L"No port with Port Owner set - EHCI already owns all");
    }

    /* Read OHCI registers for comparison (via RegBase direct access) */
    if (tryOhci && ohciRegBase) {
        DWORD testVal = 0;
        if (TryReadDword((volatile DWORD *)ohciRegBase, &testVal)) {
            DWORD rev, ctrl, rhDescA;
            int ndp;
            rev     = *(volatile DWORD *)ohciRegBase;
            ctrl    = *(volatile DWORD *)(ohciRegBase + 0x04);
            rhDescA = *(volatile DWORD *)(ohciRegBase + 0x48);
            ndp = rhDescA & 0xFF;
            wsprintfW(buf,
                L"OHCI Rev=0x%08X Ctrl=0x%08X RhDescA=0x%08X NDP=%d",
                rev, ctrl, rhDescA, ndp);
            AddLog(buf);
            for (i = 0; i < ndp && i < 4; i++) {
                DWORD rps = *(volatile DWORD *)(ohciRegBase + 0x54 + 4*i);
                wsprintfW(buf,
                    L"OHCI Port[%d]=0x%08X CCS=%u PES=%u PPS=%u LSDA=%u",
                    i, rps, rps&1, (rps>>1)&1, (rps>>8)&1, (rps>>9)&1);
                AddLog(buf);
            }
        } else {
            wsprintfW(buf, L"OHCI RegBase 0x%08X not readable", ohciRegBase);
            AddLog(buf);
        }
    }

    /* Post-check */
    AddLog(L"--- Post-check ---");
    DumpDummyClassLoadState();
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"DllMainSeen");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"UsbAttachSeen");
    {
        HANDLE hEvt; DWORD ioRc;
        hEvt = OpenEvent(EVENT_ALL_ACCESS, FALSE, L"USB_ATTACH_EVENT0");
        if (hEvt) {
            ioRc = WaitForSingleObject(hEvt, 0);
            wsprintfW(buf, L"USB_ATTACH_EVENT0: %s",
                      (ioRc == WAIT_OBJECT_0) ? L"SIGNALED" : L"not signaled");
            AddLog(buf); CloseHandle(hEvt);
        }
        hEvt = OpenEvent(EVENT_ALL_ACCESS, FALSE, L"USB_ATTACH_EVENT1");
        if (hEvt) {
            ioRc = WaitForSingleObject(hEvt, 0);
            wsprintfW(buf, L"USB_ATTACH_EVENT1: %s",
                      (ioRc == WAIT_OBJECT_0) ? L"SIGNALED" : L"not signaled");
            AddLog(buf); CloseHandle(hEvt);
        }
    }
}

static void DoEhciTakeover(void)
{
    WCHAR buf[320];
    /* USB Event names from ehci.c / ohci.c / SystemManager.c */
    static const WCHAR *evtNames[] = {
        L"USB_ATTACH_EVENT0",
        L"USB_ATTACH_EVENT1",
        L"USB_HUB_ATTACH_EVENT_D1",
        L"USB_HUB_ATTACH_EVENT_D2",
        L"USB_HUB_ATTACH_EVENT_D3",
        L"USB_HUB_ATTACH_EVENT_D4",
        L"USB_DETACH_EVENT0",
        L"USB_DETACH_EVENT1",
        L"USB_HUB_DETACH_EVENT_D1",
        L"USB_HUB_DETACH_EVENT_D2",
        L"USB_HUB_DETACH_EVENT_D3",
        L"USB_HUB_DETACH_EVENT_D4"
    };
    static const WCHAR *evtShort[] = {
        L"ATT0", L"ATT1", L"HUB_D1", L"HUB_D2", L"HUB_D3", L"HUB_D4",
        L"DET0", L"DET1", L"HDET_D1", L"HDET_D2", L"HDET_D3", L"HDET_D4"
    };
    #define N_USB_EVENTS 12
    HANDLE hEvts[N_USB_EVENTS];
    int i, iter, nOpen = 0;
    HANDLE hPio;
    unsigned char pioIn[4], pioOut[4];

    AddLog(L"=== USB Event Monitor ===");

    /* 1. Open all USB named events */
    for (i = 0; i < N_USB_EVENTS; i++) {
        hEvts[i] = OpenEvent(EVENT_ALL_ACCESS, FALSE, evtNames[i]);
        if (hEvts[i]) nOpen++;
    }
    wsprintfW(buf, L"Opened %d/%d USB events", nOpen, N_USB_EVENTS);
    AddLog(buf);

    /* Show which events exist */
    for (i = 0; i < N_USB_EVENTS; i++) {
        wsprintfW(buf, L"  %s: %s", evtNames[i],
                  hEvts[i] ? L"OK" : L"NOT FOUND");
        AddLog(buf);
    }

    if (nOpen == 0) {
        AddLog(L"No USB events found! EHCI/OHCI may not be loaded.");
        goto evt_done;
    }

    /* 2. Try PIO1: GPIO query (Panasonic USB detection GPIOs) */
    hPio = CreateFileW(L"PIO1:", 0xC0000000u, 0, NULL, 3, 0x80u, NULL);
    if (hPio && hPio != INVALID_HANDLE_VALUE) {
        AddLog(L"PIO1: opened OK (Panasonic GPIO controller)");
        /* Read GPIO pin 21 (USB root port detect from k.usbd.dll) */
        pioIn[0] = 21;
        pioOut[0] = 0xFF;
        if (DeviceIoControl(hPio, 0x1E2023u, pioIn, 1, pioOut, 1, NULL, NULL)) {
            wsprintfW(buf, L"  GPIO21 (root detect) = %u", pioOut[0]);
            AddLog(buf);
        }
        /* Read a range of GPIO pins for USB-related signals */
        {
            int pins[] = {20, 22, 23, 24, 25, 26, 27, 28, 29, 30};
            int np;
            for (np = 0; np < 10; np++) {
                pioIn[0] = (unsigned char)pins[np];
                pioOut[0] = 0xFF;
                if (DeviceIoControl(hPio, 0x1E2023u, pioIn, 1, pioOut, 1, NULL, NULL)) {
                    wsprintfW(buf, L"  GPIO%d=%u", pins[np], pioOut[0]);
                    AddLog(buf);
                }
            }
        }
        CloseHandle(hPio);
    } else {
        wsprintfW(buf, L"PIO1: open FAIL err=%u", GetLastError());
        AddLog(buf);
    }

    /* 3. Monitor USB events for 30 seconds */
    AddLog(L"--- PLUG/UNPLUG adapter NOW! Monitoring 30 sec... ---");
    for (iter = 0; iter < 60; iter++) {
        DWORD rc;
        HANDLE hWait[N_USB_EVENTS];
        int waitMap[N_USB_EVENTS];
        int nWait = 0;
        for (i = 0; i < N_USB_EVENTS; i++) {
            if (hEvts[i]) {
                hWait[nWait] = hEvts[i];
                waitMap[nWait] = i;
                nWait++;
            }
        }
        if (nWait == 0) break;
        rc = WaitForMultipleObjects(nWait, hWait, FALSE, 500);
        if (rc >= WAIT_OBJECT_0 && rc < WAIT_OBJECT_0 + (DWORD)nWait) {
            int idx = waitMap[rc - WAIT_OBJECT_0];
            wsprintfW(buf, L">>> EVENT FIRED: %s <<<", evtNames[idx]);
            AddLog(buf);
        }
    }
    AddLog(L"--- Monitor done ---");

evt_done:
    for (i = 0; i < N_USB_EVENTS; i++) {
        if (hEvts[i]) CloseHandle(hEvts[i]);
    }

    /* 4. Probe USB device handles */
    {
        static const WCHAR *devNames[] = {
            L"UHD1:", L"UHD2:", L"HCD1:", L"HCD2:",
            L"USB1:", L"USB2:", L"UHC1:", L"UHC2:"
        };
        int nd;
        for (nd = 0; nd < 8; nd++) {
            HANDLE hd = CreateFileW(devNames[nd], 0xC0000000u, 0, NULL,
                                    3, 0, NULL);
            if (hd && hd != INVALID_HANDLE_VALUE) {
                wsprintfW(buf, L"Device %s: OPEN OK", devNames[nd]);
                AddLog(buf);
                CloseHandle(hd);
            }
        }
    }
    AddLog(L"=== USB Event Monitor Done ===");
}

static void DoStageUsbDebug(void)
{
    HKEY hKey;
    DWORD disp;
    LONG rc;
    WCHAR route[] = { L'A', L'X', L'8', L'8', L'1', L'7', L'9', L'A', L'1', 0, 0 };
    WCHAR bind[] = { L'A', L'X', L'8', L'8', L'1', L'7', L'9', L'A', L'1', 0, 0 };
    DWORD busNumber = 0;
    DWORD busType = 0;
    WCHAR msg[256];

    AddLog(L"=== Direct Arm ===");

    ClearProbeState();

    if (!FileExistsW(RTL_WIN_DLL)) {
        AddLog(L"Arm FAIL: \\Windows\\rtl8152.dll is missing");
        return;
    }

    AddLog(L"Built-in driver detected: \\Windows\\rtl8152.dll");

    /* DmyClassDrv replacement: redirect the dummy class driver to our DLL.
     * Our driver accepts ALL USB devices but only inits RTL8152B
     * (VID=0x0BDA / PID=0x8152 / bcdDevice=0x2000). */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Drivers\\USB\\LoadClientsDummy\\DmyClassDrv",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: LoadClientsDummy\\DmyClassDrv -> rtl8152.dll");
    } else {
        wsprintfW(msg, L"REG FAIL: LoadClientsDummy\\DmyClassDrv rc=%d", rc);
        AddLog(msg);
    }

    /* VID_PID_bcdDevice path (k.usbd.dll tries this format FIRST) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE,
                        L"Drivers\\USB\\LoadClients\\3034_33106_8192\\Default\\Default\\RTL8152B_Class",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegSetSz(hKey, L"DLL", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: LoadClients\\3034_33106_8192 -> rtl8152.dll");
    } else {
        wsprintfW(msg, L"REG FAIL: LoadClients\\3034_33106_8192 rc=%d", rc);
        AddLog(msg);
    }

    /* VID_PID path (second fallback) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE,
                        L"Drivers\\USB\\LoadClients\\3034_33106\\Default\\Default\\RTL8152B_Class",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegSetSz(hKey, L"DLL", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: LoadClients\\3034_33106 -> rtl8152.dll");
    } else {
        wsprintfW(msg, L"REG FAIL: LoadClients\\3034_33106 rc=%d", rc);
        AddLog(msg);
    }

    /* VID_PID_bcdDevice=0x0300 (RTL8152B alternate bcd) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE,
                        L"Drivers\\USB\\LoadClients\\3034_33106_8192\\Default\\Default\\RTL8152B_Class",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegSetSz(hKey, L"DLL", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: LoadClients\\3034_33106_8192 -> rtl8152.dll");
    } else {
        wsprintfW(msg, L"REG FAIL: LoadClients\\3034_33106_8192 rc=%d", rc);
        AddLog(msg);
    }

    /* VID_PID_bcdDevice=0x0100 (RTL8152B rev A) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE,
                        L"Drivers\\USB\\LoadClients\\3034_33106_8192\\Default\\Default\\RTL8152B_Class",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegSetSz(hKey, L"DLL", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: LoadClients\\3034_33106_8192 -> rtl8152.dll");
    } else {
        wsprintfW(msg, L"REG FAIL: LoadClients\\3034_33106_8192 rc=%d", rc);
        AddLog(msg);
    }

    /* VID-only fallback (any ASIX PID) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE,
                        L"Drivers\\USB\\LoadClients\\3034\\Default\\Default\\RTL8152B_Class",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegSetSz(hKey, L"DLL", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: LoadClients\\3034 (VID-only) -> rtl8152.dll");
    } else {
        wsprintfW(msg, L"REG FAIL: LoadClients\\3034 rc=%d", rc);
        AddLog(msg);
    }

    /* Absolute catch-all: ANY USB device -> our DLL (will log VID/PID) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE,
                        L"Drivers\\USB\\LoadClients\\Default\\Default\\Default\\RTL8152B_Class",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegSetSz(hKey, L"DLL", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: LoadClients\\Default (catch-all!) -> rtl8152.dll");
    } else {
        wsprintfW(msg, L"REG FAIL: LoadClients\\Default rc=%d", rc);
        AddLog(msg);
    }

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Drivers\\USB\\ClientDrivers\\RTL8152B_Class",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"Dll", L"rtl8152.dll");
        RegSetSz(hKey, L"DLL", L"rtl8152.dll");
        RegSetSz(hKey, L"Prefix", L"RTL");
        RegCloseKey(hKey);
        AddLog(L"REG OK: ClientDrivers\\RTL8152B_Class");
    } else {
        wsprintfW(msg, L"REG FAIL: ClientDrivers\\RTL8152B_Class rc=%d", rc);
        AddLog(msg);
    }

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"DisplayName", L"Realtek RTL8152B USB Ethernet");
        RegSetSz(hKey, L"ImagePath", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: Comm\\RTL8152B");
    } else {
        wsprintfW(msg, L"REG FAIL: Comm\\RTL8152B rc=%d", rc);
        AddLog(msg);
    }

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B\\Linkage",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetValueEx(hKey, L"Route", 0, REG_MULTI_SZ, (const BYTE*)route, sizeof(route));
        RegCloseKey(hKey);
        AddLog(L"REG OK: Comm\\RTL8152B\\Linkage");
    } else {
        wsprintfW(msg, L"REG FAIL: Comm\\RTL8152B\\Linkage rc=%d", rc);
        AddLog(msg);
    }

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetSz(hKey, L"DisplayName", L"Realtek RTL8152B USB Ethernet");
        RegSetSz(hKey, L"ImagePath", L"rtl8152.dll");
        RegCloseKey(hKey);
        AddLog(L"REG OK: Comm\\RTL8152B1");
    } else {
        wsprintfW(msg, L"REG FAIL: Comm\\RTL8152B1 rc=%d", rc);
        AddLog(msg);
    }

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetDw(hKey, L"BusNumber", busNumber);
        RegSetDw(hKey, L"BusType", busType);
        RegSetDw(hKey, L"*IfType", 6);
        RegSetDw(hKey, L"*MediaType", 0);
        RegSetDw(hKey, L"*PhysicalMediaType", 0);
        RegCloseKey(hKey);
        AddLog(L"REG OK: Comm\\RTL8152B1\\Parms");
    } else {
        wsprintfW(msg, L"REG FAIL: Comm\\RTL8152B1\\Parms rc=%d", rc);
        AddLog(msg);
    }

    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\Tcpip",
                        0, NULL, 0, 0, NULL, &hKey, &disp);
    if (rc == ERROR_SUCCESS) {
        RegSetDw(hKey, L"EnableDHCP", 1);
        RegCloseKey(hKey);
        AddLog(L"REG OK: Comm\\RTL8152B1\\Parms\\Tcpip");
    } else {
        wsprintfW(msg, L"REG FAIL: Comm\\RTL8152B1\\Parms\\Tcpip rc=%d", rc);
        AddLog(msg);
    }

    (void)bind;
    AppendTcpipBindValue();

    AddLog(L"Arm OK: unplug/replug adapter or press Kick");
    AddLog(L"NOTE: Kick only drives NDIS. USBDeviceAttach requires a real replug.");
    AddLog(L"Target runtime DLL: \\Windows\\rtl8152.dll");
    AddLog(L"=== Direct Arm Done ===");
}

static void AddLogFileInfo(const WCHAR *label, const WCHAR *path)
{
    DWORD attrs;
    HANDLE hFile;
    DWORD sizeLow;
    WCHAR msg[256];

    attrs = GetFileAttributes(path);
    if (attrs == 0xFFFFFFFF) {
        wsprintfW(msg, L"%s: missing (%s)", label, path);
        AddLog(msg);
        return;
    }

    hFile = CreateFile(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_EXISTING, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE) {
        wsprintfW(msg, L"%s: exists but open FAIL err=%d", label, GetLastError());
        AddLog(msg);
        return;
    }

    sizeLow = GetFileSize(hFile, NULL);
    CloseHandle(hFile);
    wsprintfW(msg, L"%s: OK size=%u (%s)", label, sizeLow, path);
    AddLog(msg);
}

static void ProbeDeviceOpen(const WCHAR *label, const WCHAR *deviceName)
{
    HANDLE hFile;
    WCHAR msg[320];
    DWORD lastError;

    if (!deviceName || !deviceName[0]) {
        return;
    }

    hFile = CreateFile(deviceName, GENERIC_READ | GENERIC_WRITE,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_EXISTING, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE) {
        hFile = CreateFile(deviceName, GENERIC_READ,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           NULL, OPEN_EXISTING, 0, NULL);
    }

    if (hFile == INVALID_HANDLE_VALUE) {
        hFile = CreateFile(deviceName, 0,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           NULL, OPEN_EXISTING, 0, NULL);
    }

    if (hFile == INVALID_HANDLE_VALUE) {
        lastError = GetLastError();
        wsprintfW(msg, L"%s open FAIL: %s err=%d", label, deviceName, lastError);
        AddLog(msg);
        return;
    }

    wsprintfW(msg, L"%s open OK: %s", label, deviceName);
    AddLog(msg);
    CloseHandle(hFile);
}

static void ProbeUsbNamedEvents(void)
{
    static const WCHAR *eventNames[] = {
        L"USB_ATTACH_EVENT0",
        L"USB_ATTACH_EVENT1",
        L"USB_HUB_ATTACH_EVENT_D1",
        L"USB_HUB_ATTACH_EVENT_D2",
        L"USB_HUB_ATTACH_EVENT_D3",
        L"USB_DETACH_EVENT0",
        L"USB_DETACH_EVENT1",
        L"USB_HUB_DETACH_EVENT_D1",
        L"USB_HUB_DETACH_EVENT_D2",
        L"USB_HUB_DETACH_EVENT_D3",
        L"USB_OVER_CURRENT0",
        L"USB_OVER_CURRENT1",
        L"USB_OVER_CURRENT_D1",
        L"USB_OVER_CURRENT_D2",
        L"USB_OVER_CURRENT_D3",
        L"USB_CONNECT_EVENT0",
        L"USB_DISCON_EVENT0",
        L"USB_THREAD_END"
    };
    HANDLE handles[sizeof(eventNames) / sizeof(eventNames[0])];
    const WCHAR *handleNames[sizeof(eventNames) / sizeof(eventNames[0])];
    DWORD handleCount;
    DWORD i;
    DWORD waitRc;
    DWORD startTick;
    DWORD nowTick;
    WCHAR msg[320];
    int caughtAny;

    AddLog(L"--- USB named event probe ---");
    AddLog(L"Event note: SystemManager waits on the same auto-reset events, so a miss is not proof of no USB activity");

    handleCount = 0;
    for (i = 0; i < (DWORD)(sizeof(eventNames) / sizeof(eventNames[0])); ++i) {
        HANDLE hEvent;

        hEvent = OpenEventW(0x1F0003u, FALSE, eventNames[i]);
        if (!hEvent) {
            wsprintfW(msg, L"EVENT open FAIL: %s err=%d", eventNames[i], GetLastError());
            AddLog(msg);
            continue;
        }

        waitRc = WaitForSingleObject(hEvent, 0);
        if (waitRc == WAIT_OBJECT_0) {
            wsprintfW(msg, L"EVENT immediate SIGNAL: %s", eventNames[i]);
            AddLog(msg);
        } else if (waitRc == WAIT_TIMEOUT) {
            wsprintfW(msg, L"EVENT open OK: %s", eventNames[i]);
            AddLog(msg);
        } else {
            wsprintfW(msg, L"EVENT open OK but initial wait FAIL: %s rc=%u err=%d",
                      eventNames[i], waitRc, GetLastError());
            AddLog(msg);
        }

        handles[handleCount] = hEvent;
        handleNames[handleCount] = eventNames[i];
        ++handleCount;
    }

    if (handleCount == 0) {
        AddLog(L"USB event probe: no named events could be opened");
        return;
    }

    AddLog(L"USB event live wait: replug adapter now; watching 8 seconds for attach/connect/detach events");
    startTick = GetTickCount();
    caughtAny = 0;

    for (;;) {
        waitRc = WaitForMultipleObjects(handleCount, handles, FALSE, 250);
        if (waitRc >= WAIT_OBJECT_0 && waitRc < WAIT_OBJECT_0 + handleCount) {
            wsprintfW(msg, L"EVENT live SIGNAL: %s", handleNames[waitRc - WAIT_OBJECT_0]);
            AddLog(msg);
            caughtAny = 1;
            continue;
        }

        if (waitRc != WAIT_TIMEOUT) {
            wsprintfW(msg, L"EVENT live wait FAIL rc=%u err=%d", waitRc, GetLastError());
            AddLog(msg);
            break;
        }

        nowTick = GetTickCount();
        if (nowTick - startTick >= 8000u) {
            break;
        }
    }

    if (!caughtAny) {
        AddLog(L"USB event live wait: no signal caught in 8 seconds");
    }

    for (i = 0; i < handleCount; ++i) {
        CloseHandle(handles[i]);
    }
}

static void ProbePnpNotifications(void)
{
    MSGQUEUEOPTIONS msgOpts;
    HANDLE hMsgQ;
    HANDLE hNotifyUsb;
    HANDLE hNotifyNdis;
    BYTE buffer[sizeof(AX_DEVDETAIL) + 256 * sizeof(TCHAR)];
    DWORD dwFlags;
    DWORD dwRead;
    DWORD startTick;
    DWORD nowTick;
    int gotAny;
    WCHAR msg[512];

    AddLog(L"--- PnP notification probe ---");

    memset(&msgOpts, 0, sizeof(msgOpts));
    msgOpts.dwSize = sizeof(msgOpts);
    msgOpts.dwFlags = 0;
    msgOpts.dwMaxMessages = 32;
    msgOpts.cbMaxMessage = sizeof(buffer);
    msgOpts.bReadAccess = TRUE;

    hMsgQ = CreateMsgQueue(NULL, &msgOpts);
    if (!hMsgQ) {
        wsprintfW(msg, L"PnP probe FAIL: CreateMsgQueue err=%d", GetLastError());
        AddLog(msg);
        return;
    }

    hNotifyUsb = RequestDeviceNotifications(&AX_DEVCLASS_USB_DEVICE_GUID, hMsgQ, TRUE);
    hNotifyNdis = RequestDeviceNotifications(&AX_DEVCLASS_NDIS_GUID, hMsgQ, TRUE);

    if (!hNotifyUsb && !hNotifyNdis) {
        wsprintfW(msg, L"PnP probe FAIL: RequestDeviceNotifications usb=%d ndis=%d err=%d",
                  hNotifyUsb != NULL, hNotifyNdis != NULL, GetLastError());
        AddLog(msg);
        CloseHandle(hMsgQ);
        return;
    }

    wsprintfW(msg, L"PnP probe armed: USB notify=%s NDIS notify=%s",
              hNotifyUsb ? L"yes" : L"no",
              hNotifyNdis ? L"yes" : L"no");
    AddLog(msg);
    AddLog(L"PnP live wait: replug adapter now; watching 8 seconds for USB/NDIS device notifications");

    startTick = GetTickCount();
    gotAny = 0;
    for (;;) {
        if (WaitForSingleObject(hMsgQ, 250) == WAIT_OBJECT_0) {
            while (ReadMsgQueue(hMsgQ, buffer, sizeof(buffer), &dwRead, 0, &dwFlags)) {
                AX_DEVDETAIL *pDetail;

                pDetail = (AX_DEVDETAIL *)buffer;
                if (pDetail->cbName < 0) {
                    AddLog(L"PnP message skipped: invalid cbName");
                    continue;
                }
                if (pDetail->cbName >= 256 * (int)sizeof(TCHAR)) {
                    pDetail->szName[255] = L'\0';
                } else {
                    pDetail->szName[pDetail->cbName / sizeof(TCHAR)] = L'\0';
                }
                wsprintfW(msg, L"PnP %s: %s",
                          pDetail->fAttached ? L"ATTACHED" : L"DETACHED",
                          pDetail->szName);
                AddLog(msg);
                gotAny = 1;
            }
        }

        nowTick = GetTickCount();
        if (nowTick - startTick >= 8000u) {
            break;
        }
    }

    if (!gotAny) {
        AddLog(L"PnP live wait: no USB/NDIS notifications in 8 seconds");
    }

    if (hNotifyUsb) {
        StopDeviceNotifications(hNotifyUsb);
    }
    if (hNotifyNdis) {
        StopDeviceNotifications(hNotifyNdis);
    }
    CloseHandle(hMsgQ);
}

static void ProbeEhciLiveDuringWait(void)
{
    HANDLE hFile;
    DWORD last101[5];
    DWORD last103[5];
    DWORD lastRc101[5];
    DWORD lastRc103[5];
    DWORD lastErr101[5];
    DWORD lastErr103[5];
    DWORD startTick;
    DWORD nowTick;
    DWORD port;
    DWORD portValue;
    DWORD portOut101;
    DWORD portOut103;
    DWORD bytesRet;
    DWORD ioRc;
    DWORD err;
    DWORD elapsed;
    int haveState[5];
    int phase;
    WCHAR msg[320];

    hFile = CreateFile(L"HCD2:", GENERIC_READ | GENERIC_WRITE,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_EXISTING, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE) {
        wsprintfW(msg, L"EHCI live trace open FAIL: HCD2: err=%d", GetLastError());
        AddLog(msg);
        return;
    }

    memset(last101, 0, sizeof(last101));
    memset(last103, 0, sizeof(last103));
    memset(lastRc101, 0xFF, sizeof(lastRc101));
    memset(lastRc103, 0xFF, sizeof(lastRc103));
    memset(lastErr101, 0xFF, sizeof(lastErr101));
    memset(lastErr103, 0xFF, sizeof(lastErr103));
    memset(haveState, 0, sizeof(haveState));

    AddLog(L"EHCI live trace phase 1/2: unplug adapter now; watching HCD2 for 6 seconds");
    startTick = GetTickCount();
    phase = 0;
    for (;;) {
        for (port = 0; port <= 4; ++port) {
            portValue = port;
            portOut101 = 0;
            bytesRet = 0;
            SetLastError(0);
            ioRc = DeviceIoControl(hFile, 0x101u, &portValue, sizeof(BYTE),
                                   &portOut101, sizeof(portOut101), &bytesRet, NULL);
            err = GetLastError();

            if (!haveState[port] ||
                lastRc101[port] != ioRc ||
                last101[port] != portOut101 ||
                lastErr101[port] != err) {
                wsprintfW(msg,
                          L"EHCI live 0x101 port=%u %s bytes=%u value=0x%08X err=%d",
                          port,
                          ioRc ? L"OK" : L"FAIL",
                          bytesRet,
                          portOut101,
                          err);
                AddLog(msg);
                lastRc101[port] = ioRc;
                last101[port] = portOut101;
                lastErr101[port] = err;
            }

            portOut103 = 0;
            bytesRet = 0;
            SetLastError(0);
            ioRc = DeviceIoControl(hFile, 0x103u, &portValue, sizeof(BYTE),
                                   &portOut103, sizeof(portOut103), &bytesRet, NULL);
            err = GetLastError();

            if (!haveState[port] ||
                lastRc103[port] != ioRc ||
                last103[port] != portOut103 ||
                lastErr103[port] != err) {
                wsprintfW(msg,
                          L"EHCI live 0x103 port=%u %s bytes=%u value=0x%08X err=%d",
                          port,
                          ioRc ? L"OK" : L"FAIL",
                          bytesRet,
                          portOut103,
                          err);
                AddLog(msg);
                lastRc103[port] = ioRc;
                last103[port] = portOut103;
                lastErr103[port] = err;
            }

            haveState[port] = 1;
        }

        nowTick = GetTickCount();
        elapsed = nowTick - startTick;
        if (phase == 0 && elapsed >= 6000u) {
            AddLog(L"EHCI live trace phase 2/2: plug adapter now; watching HCD2 for 10 seconds");
            phase = 1;
        }
        if (elapsed >= 16000u) {
            break;
        }
        Sleep(100);
    }

    CloseHandle(hFile);
}
static void ProbeHcdIoctls(const WCHAR *label, const WCHAR *deviceName)
{
    HANDLE hFile;
    DWORD outValue;
    DWORD bytesRet;
    DWORD ioRc;
    DWORD portValue;
    DWORD portOut;
    DWORD port;
    DWORD caps[12];
    WCHAR msg[320];

    hFile = CreateFile(deviceName, GENERIC_READ | GENERIC_WRITE,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_EXISTING, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE) {
        wsprintfW(msg, L"%s IOCTL open FAIL: %s err=%d", label, deviceName, GetLastError());
        AddLog(msg);
        return;
    }

    outValue = 0;
    bytesRet = 0;
    ioRc = DeviceIoControl(hFile, 0x2A0808u, NULL, 0, &outValue, sizeof(outValue), &bytesRet, NULL);
    wsprintfW(msg, L"%s IOCTL 0x2A0808 %s bytes=%u value=0x%08X err=%d",
              label, ioRc ? L"OK" : L"FAIL", bytesRet, outValue, GetLastError());
    AddLog(msg);

    memset(caps, 0, sizeof(caps));
    bytesRet = 0;
    ioRc = DeviceIoControl(hFile, 0x321000u, NULL, 0, caps, sizeof(caps), &bytesRet, NULL);
    wsprintfW(msg,
              L"%s IOCTL 0x321000 %s bytes=%u d0=%u d1=%u d2=%u d3=%u err=%d",
              label,
              ioRc ? L"OK" : L"FAIL",
              bytesRet,
              caps[0],
              caps[1],
              caps[2],
              caps[3],
              GetLastError());
    AddLog(msg);

    outValue = 0;
    bytesRet = 0;
    ioRc = DeviceIoControl(hFile, 0x321004u, NULL, 0, &outValue, sizeof(outValue), &bytesRet, NULL);
    wsprintfW(msg, L"%s IOCTL 0x321004 %s bytes=%u value=%u (0x%08X) err=%d",
              label, ioRc ? L"OK" : L"FAIL", bytesRet, outValue, outValue, GetLastError());
    AddLog(msg);

    for (port = 0; port <= 4; ++port) {
        portValue = port;
        portOut = 0;
        bytesRet = 0;
        ioRc = DeviceIoControl(hFile, 0x101u, &portValue, sizeof(BYTE), &portOut, sizeof(portOut), &bytesRet, NULL);
        wsprintfW(msg, L"%s IOCTL 0x101 port=%u %s bytes=%u value=0x%08X err=%d",
                  label, port, ioRc ? L"OK" : L"FAIL", bytesRet, portOut, GetLastError());
        AddLog(msg);

        portOut = 0;
        bytesRet = 0;
        ioRc = DeviceIoControl(hFile, 0x103u, &portValue, sizeof(BYTE), &portOut, sizeof(portOut), &bytesRet, NULL);
        wsprintfW(msg, L"%s IOCTL 0x103 port=%u %s bytes=%u value=0x%08X err=%d",
                  label, port, ioRc ? L"OK" : L"FAIL", bytesRet, portOut, GetLastError());
        AddLog(msg);
    }

    CloseHandle(hFile);
}

static void AddRegSzValue(const WCHAR *keyPath, const WCHAR *valueName)
{
    HKEY hKey;
    WCHAR buf[260];
    DWORD cb;
    DWORD type;
    WCHAR msg[320];
    LONG rc;

    cb = sizeof(buf);
    type = 0;
    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        wsprintfW(msg, L"REG open FAIL: HKLM\\%s rc=%d", keyPath, rc);
        AddLog(msg);
        return;
    }

    rc = RegQueryValueEx(hKey, valueName, NULL, &type, (BYTE*)buf, &cb);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS || type != REG_SZ) {
        wsprintfW(msg, L"REG query FAIL: HKLM\\%s\\%s rc=%d type=%d",
                  keyPath, valueName, rc, type);
        AddLog(msg);
        return;
    }

    wsprintfW(msg, L"REG HKLM\\%s\\%s = %s", keyPath, valueName, buf);
    AddLog(msg);
}

static BOOL QueryRegSz(HKEY rootKey, const WCHAR *keyPath, const WCHAR *valueName,
                       WCHAR *buf, DWORD cchBuf)
{
    HKEY hKey;
    DWORD cb;
    DWORD type;
    LONG rc;

    if (!buf || cchBuf == 0) {
        return FALSE;
    }

    buf[0] = L'\0';
    cb = cchBuf * sizeof(WCHAR);
    type = 0;
    rc = RegOpenKeyEx(rootKey, keyPath, 0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        return FALSE;
    }

    rc = RegQueryValueEx(hKey, valueName, NULL, &type, (BYTE*)buf, &cb);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS || type != REG_SZ) {
        buf[0] = L'\0';
        return FALSE;
    }

    buf[cchBuf - 1] = L'\0';
    return TRUE;
}

static void AddRegDwValue(const WCHAR *keyPath, const WCHAR *valueName)
{
    HKEY hKey;
    DWORD value;
    DWORD cb;
    DWORD type;
    WCHAR msg[320];
    LONG rc;

    cb = sizeof(value);
    type = 0;
    value = 0;
    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        wsprintfW(msg, L"REG open FAIL: HKLM\\%s rc=%d", keyPath, rc);
        AddLog(msg);
        return;
    }

    rc = RegQueryValueEx(hKey, valueName, NULL, &type, (BYTE*)&value, &cb);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS || type != REG_DWORD || cb != sizeof(DWORD)) {
        wsprintfW(msg, L"REG query FAIL: HKLM\\%s\\%s rc=%d type=%d",
                  keyPath, valueName, rc, type);
        AddLog(msg);
        return;
    }

    wsprintfW(msg, L"REG HKLM\\%s\\%s = %u (0x%08X)",
              keyPath, valueName, value, value);
    AddLog(msg);
}

static void DumpSubkeysSimple(const WCHAR *keyPath)
{
    HKEY hKey;
    DWORD index;
    DWORD nameLen;
    WCHAR name[128];
    WCHAR msg[320];
    LONG rc;
    int found;

    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        wsprintfW(msg, L"REG open FAIL: HKLM\\%s rc=%d", keyPath, rc);
        AddLog(msg);
        return;
    }

    found = 0;
    for (index = 0; index < 16; ++index) {
        nameLen = (sizeof(name) / sizeof(name[0])) - 1;
        rc = RegEnumKeyEx(hKey, index, name, &nameLen, NULL, NULL, NULL, NULL);
        if (rc != ERROR_SUCCESS) {
            break;
        }
        name[nameLen] = L'\0';
        wsprintfW(msg, L"SUBKEY HKLM\\%s -> %s", keyPath, name);
        AddLog(msg);
        found = 1;
    }

    if (!found) {
        wsprintfW(msg, L"SUBKEY HKLM\\%s -> <none>", keyPath);
        AddLog(msg);
    }

    RegCloseKey(hKey);
}

static void DumpKeySnapshot(const WCHAR *keyPath)
{
    AddRegSzValue(keyPath, L"Key");
    AddRegSzValue(keyPath, L"Dll");
    AddRegSzValue(keyPath, L"DLL");
    AddRegSzValue(keyPath, L"Name");
    AddRegSzValue(keyPath, L"Prefix");
    AddRegSzValue(keyPath, L"BusName");
    AddRegSzValue(keyPath, L"DeviceDesc");
    AddRegSzValue(keyPath, L"FriendlyName");
    AddRegSzValue(keyPath, L"ClientDriver");
    AddRegSzValue(keyPath, L"ClientInfo");
    AddRegSzValue(keyPath, L"ActivePath");
    AddRegSzValue(keyPath, L"Parent");
    AddRegDwValue(keyPath, L"ClientInfo");
    AddRegDwValue(keyPath, L"Flags");
    AddRegDwValue(keyPath, L"Order");
    AddRegDwValue(keyPath, L"Class");
    AddRegDwValue(keyPath, L"SubClass");
    AddRegDwValue(keyPath, L"ProgIF");
    AddRegDwValue(keyPath, L"HcdCapability");
    AddRegDwValue(keyPath, L"RegBase");
    AddRegDwValue(keyPath, L"MemBase");
    AddRegDwValue(keyPath, L"CurNotify");
    AddRegDwValue(keyPath, L"CurValue");
    AddRegDwValue(keyPath, L"CurIndex");
    AddRegDwValue(keyPath, L"AndroidDetectEnable");
}

static void DumpSubkeysSnapshot(const WCHAR *keyPath, DWORD maxSubkeys)
{
    HKEY hKey;
    DWORD index;
    DWORD nameLen;
    WCHAR name[128];
    WCHAR fullPath[320];
    WCHAR msg[384];
    LONG rc;
    int found;

    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        wsprintfW(msg, L"REG open FAIL: HKLM\\%s rc=%d", keyPath, rc);
        AddLog(msg);
        return;
    }

    found = 0;
    for (index = 0; index < maxSubkeys; ++index) {
        nameLen = (sizeof(name) / sizeof(name[0])) - 1;
        rc = RegEnumKeyEx(hKey, index, name, &nameLen, NULL, NULL, NULL, NULL);
        if (rc != ERROR_SUCCESS) {
            break;
        }
        name[nameLen] = L'\0';
        wsprintfW(fullPath, L"%s\\%s", keyPath, name);
        wsprintfW(msg, L"--- Key Snapshot HKLM\\%s ---", fullPath);
        AddLog(msg);
        DumpKeySnapshot(fullPath);
        found = 1;
    }

    if (!found) {
        wsprintfW(msg, L"SUBKEY HKLM\\%s -> <none>", keyPath);
        AddLog(msg);
    }

    RegCloseKey(hKey);
}

static void ClearProbeState(void)
{
    static const WCHAR *valueNames[] = {
        L"DllMainSeen",
        L"UsbAttachSeen",
        L"DriverEntrySeen",
        L"LastStage",
        L"LastSeq",
        L"LastTick",
        L"LastVID",
        L"LastPID",
        L"LastREV",
        L"LastUniqueDriverId",
        L"LastMessage",
        L"LastLine"
    };
    HKEY hKey;
    LONG rc;
    int i;

    DeleteFile(L"\\MediaCard\\rtl8152_probe.log");

    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152BProbe", 0, 0, &hKey);
    if (rc != ERROR_SUCCESS) {
        AddLog(L"Probe state cleanup: no existing HKLM\\Comm\\RTL8152BProbe key");
        return;
    }

    for (i = 0; i < (int)(sizeof(valueNames) / sizeof(valueNames[0])); ++i) {
        RegDeleteValue(hKey, valueNames[i]);
    }
    RegCloseKey(hKey);
    AddLog(L"Probe state cleanup: cleared old RTL8152BProbe values and logs");
}

/* ===========================================================================
 * DoTestDllLoad - Build_017: Trigger EHCI Port Reset via ActivateDeviceEx
 *
 * ActivateDeviceEx asks device.exe (kernel/trusted) to load rtl8152.dll.
 * RTL_Init fires in kernel context -> DoEhciPortReset -> VirtualCopy works.
 * Forces USB 2.0 HS chirp re-negotiation for stuck USB 3.x devices.
 * =========================================================================== */
static void DoTestDllLoad(void)
{
    HMODULE hCore;
    HKEY hKey;
    DWORD val, disp;
    WCHAR msg[320];
    HANDLE hActivated = NULL;

    typedef HANDLE (WINAPI *PFN_ActivateDeviceEx)(
        LPCWSTR, LPCVOID, DWORD, LPVOID);
    typedef BOOL (WINAPI *PFN_DeactivateDevice)(HANDLE);
    PFN_ActivateDeviceEx pfnActivate = NULL;
    PFN_DeactivateDevice pfnDeactivate = NULL;

    AddLog(L"=== EHCI Port Reset (Build_017) ===");

    /* Create BuiltIn\EhciFix registry key for ActivateDeviceEx */
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Drivers\\BuiltIn\\EhciFix",
                       0, NULL, 0, 0, NULL, &hKey, &disp) == ERROR_SUCCESS) {
        RegSetValueEx(hKey, L"Dll",    0, REG_SZ,
                      (const BYTE*)L"rtl8152.dll", 28);
        RegSetValueEx(hKey, L"Prefix", 0, REG_SZ,
                      (const BYTE*)L"AXE", 8);
        val = 10;
        RegSetValueEx(hKey, L"Order",  0, REG_DWORD,
                      (const BYTE*)&val, 4);
        val = 0;
        RegSetValueEx(hKey, L"Index",  0, REG_DWORD,
                      (const BYTE*)&val, 4);
        RegCloseKey(hKey);
    } else {
        wsprintfW(msg, L"FAIL: reg err=%u", GetLastError());
        AddLog(msg);
        return;
    }

    /* Resolve ActivateDeviceEx */
    hCore = GetModuleHandle(L"coredll.dll");
    if (!hCore) hCore = LoadLibrary(L"coredll.dll");
    if (hCore) {
        pfnActivate = (PFN_ActivateDeviceEx)GetProcAddressW(
            hCore, L"ActivateDeviceEx");
        pfnDeactivate = (PFN_DeactivateDevice)GetProcAddressW(
            hCore, L"DeactivateDevice");
    }
    if (!pfnActivate) {
        AddLog(L"FAIL: ActivateDeviceEx not found");
        return;
    }

    /* Trigger: device.exe loads DLL -> RTL_Init -> DoEhciPortReset */
    AddLog(L"Calling ActivateDeviceEx -> RTL_Init -> PortReset...");
    hActivated = pfnActivate(L"Drivers\\BuiltIn\\EhciFix", NULL, 0, NULL);

    if (!hActivated) {
        wsprintfW(msg, L"ActivateDeviceEx FAIL err=%u", GetLastError());
        AddLog(msg);
    } else {
        wsprintfW(msg, L"ActivateDeviceEx OK h=0x%08X", (DWORD)hActivated);
        AddLog(msg);
        AddLog(L"Port reset fired in kernel context");
        /* Deactivate to allow re-activation later */
        Sleep(500);
        if (pfnDeactivate) pfnDeactivate(hActivated);
    }

    AddLog(L"Check driver Log for PORTSC details");
    AddLog(L"=== Port Reset Done ===");
}

static void ProbeUsbDriverLoad(void)
{
    HMODULE hCore;
    HMODULE hDriver;
    PFN_LOAD_DRIVER pLoadDriver;
    FARPROC pAttach;
    FARPROC pInstall;
    WCHAR msg[256];

    AddLog(L"--- Direct Loader Test (best effort) ---");

    hCore = GetModuleHandle(L"coredll.dll");
    if (!hCore) {
        hCore = LoadLibrary(L"coredll.dll");
    }
    if (!hCore) {
        wsprintfW(msg, L"Load test FAIL: coredll.dll unavailable err=%d", GetLastError());
        AddLog(msg);
        return;
    }

    pLoadDriver = (PFN_LOAD_DRIVER)GetProcAddress(hCore, L"LoadDriver");
    if (!pLoadDriver) {
        AddLog(L"Load test FAIL: GetProcAddress(LoadDriver) returned NULL");
        return;
    }

    hDriver = pLoadDriver(L"rtl8152.dll");
    if (!hDriver) {
        wsprintfW(msg, L"LoadDriver(rtl8152.dll) FAIL err=%d", GetLastError());
        AddLog(msg);
    } else {
        AddLog(L"LoadDriver(rtl8152.dll) OK");
        pAttach = GetProcAddress(hDriver, L"USBDeviceAttach");
        pInstall = GetProcAddress(hDriver, L"USBInstallDriver");
        wsprintfW(msg, L"Export USBDeviceAttach: %s", pAttach ? L"present" : L"missing");
        AddLog(msg);
        wsprintfW(msg, L"Export USBInstallDriver: %s", pInstall ? L"present" : L"missing");
        AddLog(msg);
        FreeLibrary(hDriver);
    }

    hDriver = LoadLibrary(L"\\Windows\\rtl8152.dll");
    if (!hDriver) {
        wsprintfW(msg, L"LoadLibrary(\\Windows\\rtl8152.dll) FAIL err=%d", GetLastError());
        AddLog(msg);
    } else {
        AddLog(L"LoadLibrary(\\Windows\\rtl8152.dll) OK");
        pAttach = GetProcAddress(hDriver, L"USBDeviceAttach");
        wsprintfW(msg, L"LoadLibrary export USBDeviceAttach: %s", pAttach ? L"present" : L"missing");
        AddLog(msg);
        FreeLibrary(hDriver);
    }
}

static void DumpDummyClassLoadState(void)
{
    HKEY hKey, hSlot;
    DWORD index, si;
    DWORD nameLen, slotNameLen;
    WCHAR name[128], slotName[128];
    WCHAR fullPath[256], slotPath[320];
    WCHAR msg[320];
    LONG rc;
    int found;

    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Drivers\\USB\\DummyClassLoad",
                      0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        AddLog(L"REG open FAIL: HKLM\\Drivers\\USB\\DummyClassLoad");
        return;
    }

    found = 0;
    for (index = 0; index < 16; ++index) {
        nameLen = (sizeof(name) / sizeof(name[0])) - 1;
        rc = RegEnumKeyEx(hKey, index, name, &nameLen, NULL, NULL, NULL, NULL);
        if (rc != ERROR_SUCCESS) {
            break;
        }
        name[nameLen] = L'\0';
        wsprintfW(fullPath, L"Drivers\\USB\\DummyClassLoad\\%s", name);
        AddLog(fullPath);
        AddRegDwValue(fullPath, L"LoadRequest");
        AddRegDwValue(fullPath, L"LoadResult");
        AddRegSzValue(fullPath, L"DriverName");

        /* Enumerate subkeys of this slot (driver entries loaded by k.usbd) */
        rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, fullPath, 0, KEY_READ, &hSlot);
        if (rc == ERROR_SUCCESS) {
            for (si = 0; si < 16; ++si) {
                slotNameLen = (sizeof(slotName) / sizeof(slotName[0])) - 1;
                rc = RegEnumKeyEx(hSlot, si, slotName, &slotNameLen, NULL, NULL, NULL, NULL);
                if (rc != ERROR_SUCCESS) break;
                slotName[slotNameLen] = L'\0';
                wsprintfW(slotPath, L"%s\\%s", fullPath, slotName);
                wsprintfW(msg, L"  subkey: %s", slotPath);
                AddLog(msg);
                AddRegSzValue(slotPath, L"Dll");
                AddRegSzValue(slotPath, L"DLL");
            }
            RegCloseKey(hSlot);
        }

        found = 1;
    }

    if (!found) {
        AddLog(L"DummyClassLoad: no slots present");
    }

    RegCloseKey(hKey);
}

static void DumpActiveDriversSummary(void)
{
    HKEY hKey;
    DWORD index;
    DWORD nameLen;
    WCHAR name[128];
    WCHAR subKey[160];
    WCHAR keyVal[260];
    WCHAR dllVal[128];
    WCHAR nameVal[128];
    WCHAR prefixVal[64];
    WCHAR msg[512];
    LONG rc;
    int found;

    AddLog(L"--- Drivers\\Active summary ---");

    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Drivers\\Active", 0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        wsprintfW(msg, L"REG open FAIL: HKLM\\Drivers\\Active rc=%d", rc);
        AddLog(msg);
        return;
    }

    found = 0;
    for (index = 0; index < 128; ++index) {
        nameLen = (sizeof(name) / sizeof(name[0])) - 1;
        rc = RegEnumKeyEx(hKey, index, name, &nameLen, NULL, NULL, NULL, NULL);
        if (rc != ERROR_SUCCESS) {
            break;
        }
        name[nameLen] = L'\0';
        wsprintfW(subKey, L"Drivers\\Active\\%s", name);

        keyVal[0] = L'\0';
        dllVal[0] = L'\0';
        nameVal[0] = L'\0';
        prefixVal[0] = L'\0';
        QueryRegSz(HKEY_LOCAL_MACHINE, subKey, L"Key", keyVal, sizeof(keyVal) / sizeof(keyVal[0]));
        QueryRegSz(HKEY_LOCAL_MACHINE, subKey, L"Dll", dllVal, sizeof(dllVal) / sizeof(dllVal[0]));
        QueryRegSz(HKEY_LOCAL_MACHINE, subKey, L"Name", nameVal, sizeof(nameVal) / sizeof(nameVal[0]));
        QueryRegSz(HKEY_LOCAL_MACHINE, subKey, L"Prefix", prefixVal, sizeof(prefixVal) / sizeof(prefixVal[0]));

        if ((keyVal[0] && (wcsstr(keyVal, L"USB") || wcsstr(keyVal, L"RTL8152B") || wcsstr(keyVal, L"DmyClassDrv") ||
                           wcsstr(keyVal, L"EHCI") || wcsstr(keyVal, L"OHCI") || wcsstr(keyVal, L"HCD"))) ||
            (dllVal[0] && (wcsstr(dllVal, L"usb") || wcsstr(dllVal, L"rtl8152") || wcsstr(dllVal, L"DmyClassDrv") ||
                           wcsstr(dllVal, L"ehci") || wcsstr(dllVal, L"ohci"))) ||
            (nameVal[0] && (wcsstr(nameVal, L"USB") || wcsstr(nameVal, L"AX") || wcsstr(nameVal, L"RTL") ||
                            wcsstr(nameVal, L"HCD"))) ||
            (prefixVal[0] && (wcsstr(prefixVal, L"USB") || wcsstr(prefixVal, L"RTL") || wcsstr(prefixVal, L"COM") ||
                              wcsstr(prefixVal, L"UKW") || wcsstr(prefixVal, L"HCD")))) {
            wsprintfW(msg, L"ACTIVE %s Key=%s Dll=%s Name=%s Prefix=%s",
                      name,
                      keyVal[0] ? keyVal : L"-",
                      dllVal[0] ? dllVal : L"-",
                      nameVal[0] ? nameVal : L"-",
                      prefixVal[0] ? prefixVal : L"-");
            AddLog(msg);

            wsprintfW(msg, L"--- Key Snapshot HKLM\\%s ---", subKey);
            AddLog(msg);
            DumpKeySnapshot(subKey);
            if (nameVal[0]) {
                ProbeDeviceOpen(L"Active Name", nameVal);
            }

            if (keyVal[0] && (wcsstr(keyVal, L"Drivers\\USB\\") || wcsstr(keyVal, L"Drivers\\Launch\\"))) {
                wsprintfW(msg, L"--- Referenced Key Snapshot HKLM\\%s ---", keyVal);
                AddLog(msg);
                DumpKeySnapshot(keyVal);
                DumpSubkeysSnapshot(keyVal, 8);
            }
            found = 1;
        }
    }

    if (!found) {
        AddLog(L"Drivers\\Active summary: no USB/AX-related active entries found");
    }

    RegCloseKey(hKey);

    AddLog(L"--- USBHCK registry snapshot ---");
    DumpKeySnapshot(L"Drivers\\USB\\USBHCK");
    DumpSubkeysSnapshot(L"Drivers\\USB\\USBHCK", 8);
    AddLog(L"--- Host controller registry snapshot ---");
    DumpKeySnapshot(L"Drivers\\USB\\USBHCU");
    DumpKeySnapshot(L"Drivers\\Launch\\OHCI");
    DumpKeySnapshot(L"Drivers\\Launch\\EHCI");
    DumpKeySnapshot(L"Drivers\\Launch\\EHCI\\SkipAndroidDetect");
    ProbeDeviceOpen(L"USB host device", L"USB1:");
    ProbeDeviceOpen(L"OHCI host device", L"HCD1:");
    ProbeDeviceOpen(L"EHCI host device", L"HCD2:");
    ProbeHcdIoctls(L"OHCI host device", L"HCD1:");
    ProbeHcdIoctls(L"EHCI host device", L"HCD2:");
}

static void ArmDummyClassLoadRequests(void)
{
    HKEY hKey;
    HKEY hSlot;
    DWORD index;
    DWORD nameLen;
    DWORD one;
    DWORD zero;
    WCHAR name[128];
    WCHAR fullPath[256];
    WCHAR msg[320];
    LONG rc;
    int found;

    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Drivers\\USB\\DummyClassLoad",
                      0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        AddLog(L"DummyClassLoad arm: no slots present yet");
        return;
    }

    found = 0;
    one = 1;
    zero = 0;
    for (index = 0; index < 16; ++index) {
        nameLen = (sizeof(name) / sizeof(name[0])) - 1;
        rc = RegEnumKeyEx(hKey, index, name, &nameLen, NULL, NULL, NULL, NULL);
        if (rc != ERROR_SUCCESS) {
            break;
        }
        name[nameLen] = L'\0';
        wsprintfW(fullPath, L"Drivers\\USB\\DummyClassLoad\\%s", name);
        rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, fullPath, 0, NULL, 0, 0,
                            NULL, &hSlot, NULL);
        if (rc == ERROR_SUCCESS) {
            RegSetValueEx(hSlot, L"LoadRequest", 0, REG_DWORD,
                          (const BYTE*)&one, sizeof(one));
            RegSetValueEx(hSlot, L"LoadResult", 0, REG_DWORD,
                          (const BYTE*)&zero, sizeof(zero));
            RegCloseKey(hSlot);
            wsprintfW(msg, L"DummyClassLoad armed: HKLM\\%s LoadRequest=1 LoadResult=0",
                      fullPath);
            AddLog(msg);
        } else {
            wsprintfW(msg, L"DummyClassLoad arm FAIL: HKLM\\%s rc=%d", fullPath, rc);
            AddLog(msg);
        }
        found = 1;
    }

    if (!found) {
        AddLog(L"DummyClassLoad arm: no slots present yet");
    }

    RegCloseKey(hKey);
}

static void DoProbeDriver(void)
{
    AddLog(L"=== Driver Probe ===");
    AddRegSzValue(L"Comm\\RTL8152B", L"ImagePath");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\RTL8152B_Class", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\RTL8152B_Class", L"DLL");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\RTL8152B_Class", L"Prefix");
    ProbeAxLoadClientBase(L"Drivers\\USB\\LoadClients\\3034_33106_8192");
    ProbeAxLoadClientBase(L"Drivers\\USB\\LoadClients\\3034_33106");
    ProbeAxLoadClientBase(L"Drivers\\USB\\LoadClients\\3034_33106");
    ProbeAxLoadClientBase(L"Drivers\\USB\\LoadClients\\3034");
    AddRegDwValue(L"Drivers\\Launch\\EHCI\\SkipAndroidDetect", L"RTL8152B_VIDPID");
    ProbeAndroidDriverRoute();

    AddLog(L"--- Panasonic stock USB client reference ---");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\USBDSRC_CLASS", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\USBDSRC_CLASS", L"Prefix");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\USBEXBOX_CLASS", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\USBEXBOX_CLASS", L"Prefix");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\USBLUF_CLASS", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\ClientDrivers\\USBLUF_CLASS", L"Prefix");
    AddRegSzValue(L"Drivers\\USB\\LoadClients\\1242_12802_256\\Default\\Default\\USBDSRC_CLASS", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\LoadClients\\1242_12812_256\\Default\\Default\\USBEXBOX_CLASS", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\LoadClients\\1242_32_256\\Default\\Default\\USBEXBOX_CLASS", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\LoadClients\\Default\\2_0_0\\Default\\USBLUF_CLASS", L"Dll");
    AddRegSzValue(L"Drivers\\USB\\LoadClientsDummy\\DmyClassDrv", L"Dll");

    DumpSubkeysSimple(L"Drivers\\USB\\LoadClients\\1242_12802_256\\Default\\Default");
    DumpSubkeysSimple(L"Drivers\\USB\\LoadClients\\1242_12812_256\\Default\\Default");
    DumpSubkeysSimple(L"Drivers\\USB\\LoadClients\\1242_32_256\\Default\\Default");
    DumpSubkeysSimple(L"Drivers\\USB\\LoadClients\\Default\\2_0_0\\Default");
    DumpSubkeysSimple(L"Drivers\\USB\\LoadClientsDummy");
    AddLogFileInfo(L"DLL in \\Windows", L"\\Windows\\rtl8152.dll");
    AddLogFileInfo(L"Probe log in \\Windows", L"\\Windows\\rtl8152_probe.log");
    AddLogFileInfo(L"Probe log in \\MediaCard", L"\\MediaCard\\rtl8152_probe.log");
    AddLogFileInfo(L"RAM DLL in \\Windows", L"\\Windows\\rtl8152_ram.dll");
    AddLogFileInfo(L"NDIS in \\Windows", L"\\Windows\\NDIS.dll");
    ProbeDeviceOpen(L"USB disk-style path", L"\\USB_Dev");
    DumpDummyClassLoadState();
    ProbeUsbNamedEvents();
    ProbePnpNotifications();
    AddLog(L"--- EHCI live port trace ---");
    AddLog(L"EHCI live trace: follow prompts exactly: first unplug, then plug, while raw HCD2 state is sampled");
    ProbeEhciLiveDuringWait();
    AddLog(L"--- DummyClassLoad state after live waits ---");
    DumpDummyClassLoadState();
    DumpActiveDriversSummary();
    AddLog(L"--- Probe Registry HKLM\\Comm\\RTL8152BProbe ---");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"DllMainSeen");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"UsbAttachSeen");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"DriverEntrySeen");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"LastStage");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"LastSeq");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"LastTick");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"LastVID");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"LastPID");
    AddRegDwValue(L"Comm\\RTL8152BProbe", L"LastREV");
    AddRegSzValue(L"Comm\\RTL8152BProbe", L"LastUniqueDriverId");
    AddRegSzValue(L"Comm\\RTL8152BProbe", L"LastMessage");
    AddRegSzValue(L"Comm\\RTL8152BProbe", L"LastLine");
    ProbeUsbDriverLoad();
    AddLog(L"Probe hint: after a real replug, LoadClients should match 3034_33106_8192 first for RTL8152B REV 0200");
    AddLog(L"Probe interpretation: if RTL8152BProbe key is missing entirely, the probe DLL never loaded");
    AddLog(L"Probe interpretation: DllMainSeen=1 with UsbAttachSeen missing means DLL loaded but USB attach did not fire");
    AddLog(L"Direct loader test is best-effort only; real USB attach still requires a physical replug");
    AddLog(L"=== Probe Done ===");
}

/* ---- ICMP (loaded at runtime from iphlpapi.dll) ---- */
typedef HANDLE (WINAPI *PFN_IcmpCreate)(void);
typedef BOOL   (WINAPI *PFN_IcmpClose)(HANDLE);
typedef DWORD  (WINAPI *PFN_IcmpSend)(HANDLE, DWORD, LPVOID, WORD,
                                      LPVOID, LPVOID, DWORD, DWORD);

/* ---- ICMP Ping Server ---- */
#define AX_ICMP_ECHO_REQ  8
#define AX_ICMP_ECHO_REP  0

#pragma pack(push, 1)
typedef struct {
    BYTE   type;
    BYTE   code;
    USHORT cksum;
    USHORT id;
    USHORT seq;
} AX_ICMP_HDR;
#pragma pack(pop)

static volatile int g_bPingSrvRun = 0;
static SOCKET       g_sPingSrv = INVALID_SOCKET;

static USHORT IcmpCksum(const BYTE *data, int len)
{
    ULONG sum = 0;
    const USHORT *p = (const USHORT *)data;
    while (len > 1) { sum += *p++; len -= 2; }
    if (len == 1) sum += *(const BYTE *)p;
    sum = (sum >> 16) + (sum & 0xFFFF);
    sum += (sum >> 16);
    return (USHORT)(~sum);
}

/* ============================================================================
 * AddLog - append to listbox + log file
 * ============================================================================ */
static void AddLog(const WCHAR *msg)
{
    HANDLE hFile;
    DWORD written;
    WCHAR buf[512];
    SYSTEMTIME st;
    char abuf[1024];
    int i;
    LRESULT idx;

    GetLocalTime(&st);
    wsprintfW(buf, L"[%02d:%02d:%02d] %s",
              st.wHour, st.wMinute, st.wSecond, msg);

    if (g_hList) {
        idx = SendMessage(g_hList, LB_ADDSTRING, 0, (LPARAM)buf);
        SendMessage(g_hList, LB_SETTOPINDEX, idx, 0);
        UpdateWindow(g_hList);
    }

    hFile = CreateFile(LOG_FILE, GENERIC_WRITE, FILE_SHARE_READ,
                       NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile != INVALID_HANDLE_VALUE) {
        SetFilePointer(hFile, 0, NULL, FILE_END);
        i = 0;
        while (buf[i] && i < 1020) { abuf[i] = (char)buf[i]; i++; }
        abuf[i] = '\r'; abuf[i + 1] = '\n';
        WriteFile(hFile, abuf, i + 2, &written, NULL);
        CloseHandle(hFile);
    }
}

/* ============================================================================
 * Registry helpers
 * ============================================================================ */
static void RegSetSz(HKEY hKey, const WCHAR *name, const WCHAR *val)
{
    RegSetValueEx(hKey, name, 0, REG_SZ,
                  (const BYTE*)val,
                  (DWORD)((wcslen(val) + 1) * sizeof(WCHAR)));
}

/* Write a single string as REG_MULTI_SZ (double-null terminated) */
static void RegSetMsz1(HKEY hKey, const WCHAR *name, const WCHAR *val)
{
    WCHAR buf[128];
    DWORD len = (DWORD)wcslen(val);
    memset(buf, 0, sizeof(buf));
    wcscpy(buf, val);
    /* buf[len] = '\0' from wcscpy, buf[len+1] = '\0' from memset */
    RegSetValueEx(hKey, name, 0, REG_MULTI_SZ,
                  (const BYTE*)buf,
                  (DWORD)((len + 2) * sizeof(WCHAR)));
}

static void RegSetDw(HKEY hKey, const WCHAR *name, DWORD val)
{
    RegSetValueEx(hKey, name, 0, REG_DWORD,
                  (const BYTE*)&val, sizeof(DWORD));
}

/* ============================================================================
 * ParseIPA - narrow "a.b.c.d" -> DWORD (network byte order)
 * ============================================================================ */
static DWORD ParseIPA(const char *s)
{
    unsigned p[4];
    int pi, i;
    p[0] = p[1] = p[2] = p[3] = 0;
    for (i = 0, pi = 0; s[i] && pi < 4; i++) {
        if (s[i] >= '0' && s[i] <= '9')
            p[pi] = p[pi] * 10 + (unsigned)(s[i] - '0');
        else if (s[i] == '.')
            pi++;
    }
    return p[0] | (p[1] << 8) | (p[2] << 16) | (p[3] << 24);
}

/* ============================================================================
 * ParseIPA2 - wide L"a.b.c.d" -> DWORD (network byte order)
 * ============================================================================ */
static DWORD ParseIPA2(const WCHAR *s)
{
    unsigned p[4];
    int pi, i;
    p[0] = p[1] = p[2] = p[3] = 0;
    for (i = 0, pi = 0; s[i] && pi < 4; i++) {
        if (s[i] >= L'0' && s[i] <= L'9')
            p[pi] = p[pi] * 10 + (unsigned)(s[i] - L'0');
        else if (s[i] == L'.')
            pi++;
    }
    return p[0] | (p[1] << 8) | (p[2] << 16) | (p[3] << 24);
}

/* ============================================================================
 * FormatIP - DWORD (NBO) -> L"a.b.c.d"
 * ============================================================================ */
static void FormatIP(DWORD ip, WCHAR *buf)
{
    wsprintfW(buf, L"%d.%d.%d.%d",
              ip & 0xFF, (ip >> 8) & 0xFF,
              (ip >> 16) & 0xFF, (ip >> 24) & 0xFF);
}

static BOOL ResolveHostA(const char *host, DWORD *ipNBO)
{
    DWORD ip;
    struct hostent *he;

    if (!host || !ipNBO)
        return FALSE;

    ip = inet_addr(host);
    if (ip != INADDR_NONE) {
        *ipNBO = ip;
        return TRUE;
    }

    he = gethostbyname(host);
    if (!he || he->h_addrtype != AF_INET || !he->h_addr_list || !he->h_addr_list[0])
        return FALSE;

    memcpy(ipNBO, he->h_addr_list[0], 4);
    return TRUE;
}

static BOOL ConnectTcpHost(const char *host, USHORT port, DWORD timeoutMs,
                           SOCKET *outSock, DWORD *outIpNBO)
{
    SOCKET s;
    DWORD ipNBO;
    struct sockaddr_in sa;
    u_long nonBlock;
    fd_set wfds;
    fd_set efds;
    TIMEVAL tv;
    int rc;
    int soErr;
    int soLen;
    DWORD timeo;

    if (!outSock)
        return FALSE;

    *outSock = INVALID_SOCKET;
    if (outIpNBO)
        *outIpNBO = 0;

    if (!ResolveHostA(host, &ipNBO))
        return FALSE;

    s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (s == INVALID_SOCKET)
        return FALSE;

    nonBlock = 1;
    ioctlsocket(s, FIONBIO, &nonBlock);

    memset(&sa, 0, sizeof(sa));
    sa.sin_family = AF_INET;
    sa.sin_port = htons(port);
    sa.sin_addr.s_addr = ipNBO;

    rc = connect(s, (struct sockaddr*)&sa, sizeof(sa));
    if (rc != 0) {
        int se = WSAGetLastError();
        if (se != WSAEWOULDBLOCK && se != WSAEINPROGRESS && se != WSAEINVAL) {
            closesocket(s);
            return FALSE;
        }

        FD_ZERO(&wfds);
        FD_ZERO(&efds);
        FD_SET(s, &wfds);
        FD_SET(s, &efds);
        tv.tv_sec = timeoutMs / 1000;
        tv.tv_usec = (timeoutMs % 1000) * 1000;

        rc = select(0, NULL, &wfds, &efds, &tv);
        if (rc <= 0) {
            closesocket(s);
            return FALSE;
        }

        soErr = 0;
        soLen = sizeof(soErr);
        if (getsockopt(s, SOL_SOCKET, SO_ERROR, (char*)&soErr, &soLen) != 0 || soErr != 0) {
            closesocket(s);
            return FALSE;
        }
    }

    nonBlock = 0;
    ioctlsocket(s, FIONBIO, &nonBlock);

    timeo = timeoutMs;
    setsockopt(s, SOL_SOCKET, SO_RCVTIMEO, (const char*)&timeo, sizeof(timeo));
    setsockopt(s, SOL_SOCKET, SO_SNDTIMEO, (const char*)&timeo, sizeof(timeo));

    *outSock = s;
    if (outIpNBO)
        *outIpNBO = ipNBO;
    return TRUE;
}

static BOOL HttpReadHeaders(SOCKET s, char *hdrBuf, int hdrBufSize,
                            int *statusCode, int *bodyOffset, int *bytesInBuffer)
{
    int total;
    int n;
    int i;

    if (!hdrBuf || hdrBufSize < 64)
        return FALSE;

    total = 0;
    hdrBuf[0] = '\0';
    if (statusCode) *statusCode = 0;
    if (bodyOffset) *bodyOffset = 0;
    if (bytesInBuffer) *bytesInBuffer = 0;

    while (total < hdrBufSize - 1) {
        n = recv(s, hdrBuf + total, hdrBufSize - 1 - total, 0);
        if (n <= 0)
            return FALSE;
        total += n;
        hdrBuf[total] = '\0';

        for (i = 3; i < total; i++) {
            if (hdrBuf[i - 3] == '\r' && hdrBuf[i - 2] == '\n' &&
                hdrBuf[i - 1] == '\r' && hdrBuf[i] == '\n') {
                if (bodyOffset) *bodyOffset = i + 1;
                if (bytesInBuffer) *bytesInBuffer = total;
                if (statusCode) {
                    int code = 0;
                    sscanf(hdrBuf, "HTTP/%*d.%*d %d", &code);
                    *statusCode = code;
                }
                return TRUE;
            }
        }
    }

    return FALSE;
}

static BOOL HttpSimpleGet(const char *host, const char *path,
                          DWORD *statusCode, DWORD *elapsedMs, DWORD *bodyBytes)
{
    SOCKET s;
    char req[512];
    char hdr[2048];
    DWORD startTick;
    int code;
    int bodyOffset;
    int bytesInBuffer;
    DWORD totalBody;
    int n;

    if (statusCode) *statusCode = 0;
    if (elapsedMs) *elapsedMs = 0;
    if (bodyBytes) *bodyBytes = 0;

    if (!ConnectTcpHost(host, 80, 5000, &s, NULL))
        return FALSE;

    req[0] = '\0';
    strcat(req, "GET ");
    strcat(req, path);
    strcat(req, " HTTP/1.0\r\nHost: ");
    strcat(req, host);
    strcat(req, "\r\nUser-Agent: NetConfig/1.0\r\nConnection: close\r\n\r\n");

    startTick = GetTickCount();
    if (send(s, req, (int)strlen(req), 0) <= 0) {
        closesocket(s);
        return FALSE;
    }

    if (!HttpReadHeaders(s, hdr, sizeof(hdr), &code, &bodyOffset, &bytesInBuffer)) {
        closesocket(s);
        return FALSE;
    }

    totalBody = 0;
    if (bytesInBuffer > bodyOffset)
        totalBody += (DWORD)(bytesInBuffer - bodyOffset);

    for (;;) {
        char buf[1024];
        n = recv(s, buf, sizeof(buf), 0);
        if (n <= 0)
            break;
        totalBody += (DWORD)n;
    }

    closesocket(s);

    if (statusCode) *statusCode = (DWORD)code;
    if (elapsedMs) *elapsedMs = GetTickCount() - startTick;
    if (bodyBytes) *bodyBytes = totalBody;
    return TRUE;
}

static BOOL HttpDownloadForMs(const char *host, const char *path, DWORD durationMs,
                              DWORD *elapsedMs, DWORD *bodyBytes, DWORD *statusCode)
{
    SOCKET s;
    char req[512];
    char hdr[2048];
    DWORD startTick;
    int code;
    int bodyOffset;
    int bytesInBuffer;
    DWORD totalBody;

    if (elapsedMs) *elapsedMs = 0;
    if (bodyBytes) *bodyBytes = 0;
    if (statusCode) *statusCode = 0;

    if (!ConnectTcpHost(host, 80, 5000, &s, NULL))
        return FALSE;

    req[0] = '\0';
    strcat(req, "GET ");
    strcat(req, path);
    strcat(req, " HTTP/1.0\r\nHost: ");
    strcat(req, host);
    strcat(req, "\r\nUser-Agent: NetConfig/1.0\r\nConnection: close\r\n\r\n");

    if (send(s, req, (int)strlen(req), 0) <= 0) {
        closesocket(s);
        return FALSE;
    }

    if (!HttpReadHeaders(s, hdr, sizeof(hdr), &code, &bodyOffset, &bytesInBuffer)) {
        closesocket(s);
        return FALSE;
    }

    totalBody = 0;
    if (bytesInBuffer > bodyOffset)
        totalBody += (DWORD)(bytesInBuffer - bodyOffset);

    startTick = GetTickCount();
    for (;;) {
        char buf[2048];
        int n;
        DWORD now;

        now = GetTickCount();
        if (now - startTick >= durationMs)
            break;

        n = recv(s, buf, sizeof(buf), 0);
        if (n <= 0)
            break;
        totalBody += (DWORD)n;
    }

    closesocket(s);

    if (elapsedMs) *elapsedMs = GetTickCount() - startTick;
    if (bodyBytes) *bodyBytes = totalBody;
    if (statusCode) *statusCode = (DWORD)code;
    return TRUE;
}

static void LogResolvedHost(const char *host)
{
    DWORD ipNBO;
    DWORD startTick;
    WCHAR whost[96];
    WCHAR wip[32];
    WCHAR msg[160];

    AtoW(host, whost, 96);
    startTick = GetTickCount();
    if (ResolveHostA(host, &ipNBO)) {
        FormatIP(ipNBO, wip);
        wsprintfW(msg, L"  %s -> %s (%u ms)", whost, wip, GetTickCount() - startTick);
    } else {
        wsprintfW(msg, L"  %s -> resolve FAIL err=%d", whost, WSAGetLastError());
    }
    AddLog(msg);
}

/* ============================================================================
 * AtoW - narrow string to wide (dst must hold maxLen WCHARs)
 * ============================================================================ */
static void AtoW(const char *src, WCHAR *dst, int maxLen)
{
    int i;
    for (i = 0; src[i] && i < maxLen - 1; i++)
        dst[i] = (WCHAR)(unsigned char)src[i];
    dst[i] = 0;
}

static BOOL QueryRegMultiSzFirst(HKEY rootKey, const WCHAR *keyPath,
                                 const WCHAR *valueName, WCHAR *buf, DWORD cchBuf)
{
    HKEY hKey;
    DWORD cb;
    DWORD type;
    LONG rc;

    if (!buf || cchBuf == 0) {
        return FALSE;
    }

    buf[0] = L'\0';
    cb = cchBuf * sizeof(WCHAR);
    type = 0;

    rc = RegOpenKeyEx(rootKey, keyPath, 0, KEY_READ, &hKey);
    if (rc != ERROR_SUCCESS) {
        return FALSE;
    }

    rc = RegQueryValueEx(hKey, valueName, NULL, &type, (BYTE*)buf, &cb);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS) {
        buf[0] = L'\0';
        return FALSE;
    }

    if (type != REG_MULTI_SZ && type != REG_SZ) {
        buf[0] = L'\0';
        return FALSE;
    }

    buf[cchBuf - 1] = L'\0';
    return (buf[0] != L'\0');
}

static BOOL ParseIpv4W(const WCHAR *src, int octets[4])
{
    int part;
    int value;

    if (!src || !octets) {
        return FALSE;
    }

    for (part = 0; part < 4; part++) {
        value = 0;
        if (*src < L'0' || *src > L'9') {
            return FALSE;
        }

        while (*src >= L'0' && *src <= L'9') {
            value = value * 10 + (*src - L'0');
            if (value > 255) {
                return FALSE;
            }
            src++;
        }

        octets[part] = value;
        if (part < 3) {
            if (*src != L'.') {
                return FALSE;
            }
            src++;
        }
    }

    return (*src == L'\0');
}

/* ============================================================================
 * Static IP Config Dialog - touchscreen-friendly editor
 *
 * 4 rows: IP, Mask, GW, DNS1.  Each row = 4 octets (0..255).
 * Tap an octet to select it.  Use on-screen arrows to change value.
 * OK applies, Cancel discards.
 * ============================================================================ */

#define IPCFG_FIELDS  4    /* IP, Mask, GW, DNS1 */
#define IPCFG_OCTETS  4

/* Dialog result: 1 = OK, 0 = cancelled */
static int           g_cfgResult;
static int           g_cfg[IPCFG_FIELDS][IPCFG_OCTETS];
static int           g_selField;   /* 0-3 = IP/Mask/GW/DNS */
static int           g_selOctet;   /* 0-3 */
static HWND          g_hCfgWnd;
static const WCHAR  *g_cfgLabels[IPCFG_FIELDS] = {
    L"IP  :", L"Mask:", L"GW  :", L"DNS :"
};

static void LoadStaticCfgFromRegistry(void)
{
    WCHAR value[64];

    if (QueryRegMultiSzFirst(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\Tcpip",
                             L"IpAddress", value,
                             sizeof(value) / sizeof(value[0]))) {
        ParseIpv4W(value, g_cfg[0]);
    }

    if (QueryRegMultiSzFirst(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\Tcpip",
                             L"SubnetMask", value,
                             sizeof(value) / sizeof(value[0]))) {
        ParseIpv4W(value, g_cfg[1]);
    }

    if (QueryRegMultiSzFirst(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\Tcpip",
                             L"DefaultGateway", value,
                             sizeof(value) / sizeof(value[0]))) {
        ParseIpv4W(value, g_cfg[2]);
    }

    if (QueryRegMultiSzFirst(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\Tcpip",
                             L"DNS", value,
                             sizeof(value) / sizeof(value[0]))) {
        ParseIpv4W(value, g_cfg[3]);
    }
}

/* Button IDs inside config dialog */
#define IDC_CFG_UP     200
#define IDC_CFG_DN     201
#define IDC_CFG_LEFT   202
#define IDC_CFG_RIGHT  203
#define IDC_CFG_OK     204
#define IDC_CFG_CANCEL 205
#define IDC_CFG_P10    206
#define IDC_CFG_M10    207

static void CfgDefaults(void)
{
    /* IP: 192.168.50.100 */
    g_cfg[0][0] = 192; g_cfg[0][1] = 168; g_cfg[0][2] = 50; g_cfg[0][3] = 100;
    /* Mask: 255.255.255.0 */
    g_cfg[1][0] = 255; g_cfg[1][1] = 255; g_cfg[1][2] = 255; g_cfg[1][3] = 0;
    /* GW: 192.168.50.1 */
    g_cfg[2][0] = 192; g_cfg[2][1] = 168; g_cfg[2][2] = 50; g_cfg[2][3] = 1;
    /* DNS: 192.168.50.1 */
    g_cfg[3][0] = 192; g_cfg[3][1] = 168; g_cfg[3][2] = 50; g_cfg[3][3] = 1;

    /* Prefer values already stored in registry when available. */
    LoadStaticCfgFromRegistry();

    g_selField = 0;
    g_selOctet = 0;
}

static void CfgToStr(int field, WCHAR *buf)
{
    wsprintfW(buf, L"%d.%d.%d.%d",
              g_cfg[field][0], g_cfg[field][1],
              g_cfg[field][2], g_cfg[field][3]);
}

/* Paint the config dialog contents */
static void CfgPaint(HWND hWnd)
{
    PAINTSTRUCT ps;
    HDC hdc;
    int f, o, x, y;
    RECT rc, cell;
    HBRUSH brSel, brNorm, brWhite;
    HFONT hFont, hOld;
    WCHAR tmp[8];

    hdc = BeginPaint(hWnd, &ps);
    GetClientRect(hWnd, &rc);

    hFont = (HFONT)GetStockObject(SYSTEM_FONT);
    hOld = (HFONT)SelectObject(hdc, hFont);

    brSel   = CreateSolidBrush(RGB(0, 120, 215));  /* blue */
    brNorm  = CreateSolidBrush(RGB(230, 230, 230));
    brWhite = CreateSolidBrush(RGB(255, 255, 255));

    /* Background */
    FillRect(hdc, &rc, brWhite);

    /* Title */
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, RGB(0, 0, 0));
    {
        RECT tr = {10, 5, rc.right - 10, 30};
        DrawText(hdc, L"Static IP Config (Tap octet / use arrows)",
                 -1, &tr, DT_LEFT | DT_SINGLELINE | DT_VCENTER);
    }

    /* 4 rows: label + 4 octet cells */
    for (f = 0; f < IPCFG_FIELDS; f++) {
        y = 35 + f * 42;

        /* Label */
        {
            RECT lr = {10, y, 60, y + 36};
            DrawText(hdc, g_cfgLabels[f], -1, &lr,
                     DT_LEFT | DT_SINGLELINE | DT_VCENTER);
        }

        /* 4 octets */
        for (o = 0; o < IPCFG_OCTETS; o++) {
            x = 65 + o * 70;
            cell.left   = x;
            cell.top    = y;
            cell.right  = x + 60;
            cell.bottom = y + 36;

            if (f == g_selField && o == g_selOctet) {
                FillRect(hdc, &cell, brSel);
                SetTextColor(hdc, RGB(255, 255, 255));
            } else {
                FillRect(hdc, &cell, brNorm);
                SetTextColor(hdc, RGB(0, 0, 0));
            }

            /* Border */
            {
                HPEN hPen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
                HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
                HBRUSH hOldBr = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
                Rectangle(hdc, cell.left, cell.top, cell.right, cell.bottom);
                SelectObject(hdc, hOldBr);
                SelectObject(hdc, hOldPen);
                DeleteObject(hPen);
            }

            wsprintfW(tmp, L"%d", g_cfg[f][o]);
            DrawText(hdc, tmp, -1, &cell,
                     DT_CENTER | DT_SINGLELINE | DT_VCENTER);
        }

        /* Dots between octets */
        SetTextColor(hdc, RGB(0, 0, 0));
        for (o = 0; o < 3; o++) {
            RECT dr;
            dr.left = 65 + (o + 1) * 70 - 8;
            dr.top = y;
            dr.right = dr.left + 8;
            dr.bottom = y + 36;
            DrawText(hdc, L".", 1, &dr,
                     DT_CENTER | DT_SINGLELINE | DT_VCENTER);
        }
    }

    SelectObject(hdc, hOld);
    DeleteObject(brSel);
    DeleteObject(brNorm);
    DeleteObject(brWhite);
    EndPaint(hWnd, &ps);
}

/* Handle tap on an octet cell */
static void CfgHitTest(int mx, int my)
{
    int f, o, x, y;
    for (f = 0; f < IPCFG_FIELDS; f++) {
        y = 35 + f * 42;
        if (my < y || my > y + 36) continue;
        for (o = 0; o < IPCFG_OCTETS; o++) {
            x = 65 + o * 70;
            if (mx >= x && mx <= x + 60) {
                g_selField = f;
                g_selOctet = o;
                InvalidateRect(g_hCfgWnd, NULL, FALSE);
                return;
            }
        }
    }
}

static void CfgAdjust(int delta)
{
    int v = g_cfg[g_selField][g_selOctet] + delta;
    if (v > 255) v = 255;
    if (v < 0)   v = 0;
    g_cfg[g_selField][g_selOctet] = v;
    InvalidateRect(g_hCfgWnd, NULL, FALSE);
}

static void CfgMove(int df, int do2)
{
    g_selField += df;
    g_selOctet += do2;
    if (g_selField < 0) g_selField = IPCFG_FIELDS - 1;
    if (g_selField >= IPCFG_FIELDS) g_selField = 0;
    if (g_selOctet < 0) g_selOctet = IPCFG_OCTETS - 1;
    if (g_selOctet >= IPCFG_OCTETS) g_selOctet = 0;
    InvalidateRect(g_hCfgWnd, NULL, FALSE);
}

static LRESULT CALLBACK CfgWndProc(HWND hWnd, UINT msg,
                                    WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_PAINT:
        CfgPaint(hWnd);
        return 0;

    case WM_LBUTTONDOWN:
        CfgHitTest((int)(short)LOWORD(lParam), (int)(short)HIWORD(lParam));
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_CFG_UP:    CfgAdjust(+1);    break;
        case IDC_CFG_DN:    CfgAdjust(-1);    break;
        case IDC_CFG_P10:   CfgAdjust(+10);   break;
        case IDC_CFG_M10:   CfgAdjust(-10);   break;
        case IDC_CFG_LEFT:  CfgMove(0, -1);   break;
        case IDC_CFG_RIGHT: CfgMove(0, +1);   break;
        case IDC_CFG_OK:
            g_cfgResult = 1;
            DestroyWindow(hWnd);
            break;
        case IDC_CFG_CANCEL:
            g_cfgResult = 0;
            DestroyWindow(hWnd);
            break;
        }
        return 0;

    case WM_DESTROY:
        g_hCfgWnd = NULL;
        PostQuitMessage(0);
        return 0;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
}

/* Show the config dialog, block until closed. Returns 1 if OK. */
static int ShowStaticConfigDialog(HINSTANCE hInst)
{
    WNDCLASS wc;
    MSG dmsg;
    int bw, bh, bx, by;

    CfgDefaults();
    g_cfgResult = 0;

    memset(&wc, 0, sizeof(wc));
    wc.lpfnWndProc = CfgWndProc;
    wc.hInstance = hInst;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"AX_IPCfg";
    RegisterClass(&wc);

    g_hCfgWnd = CreateWindow(L"AX_IPCfg", L"Static IP Settings",
        WS_POPUP | WS_VISIBLE | WS_CAPTION | WS_SYSMENU,
        5, 5, 440, 310, NULL, NULL, hInst, NULL);

    if (!g_hCfgWnd) return 0;

    /* Arrow buttons below the fields */
    bw = 55; bh = 36;
    by = 210;

    bx = 15;
    CreateWindow(L"BUTTON", L"\x25C0",   /* LEFT arrow */
        WS_CHILD | WS_VISIBLE, bx, by, bw, bh,
        g_hCfgWnd, (HMENU)IDC_CFG_LEFT, NULL, NULL);
    bx += bw + 5;

    CreateWindow(L"BUTTON", L"\x25B6",   /* RIGHT arrow */
        WS_CHILD | WS_VISIBLE, bx, by, bw, bh,
        g_hCfgWnd, (HMENU)IDC_CFG_RIGHT, NULL, NULL);
    bx += bw + 20;

    CreateWindow(L"BUTTON", L"-10",
        WS_CHILD | WS_VISIBLE, bx, by, bw, bh,
        g_hCfgWnd, (HMENU)IDC_CFG_M10, NULL, NULL);
    bx += bw + 5;

    CreateWindow(L"BUTTON", L"-1",
        WS_CHILD | WS_VISIBLE, bx, by, bw, bh,
        g_hCfgWnd, (HMENU)IDC_CFG_DN, NULL, NULL);
    bx += bw + 5;

    CreateWindow(L"BUTTON", L"+1",
        WS_CHILD | WS_VISIBLE, bx, by, bw, bh,
        g_hCfgWnd, (HMENU)IDC_CFG_UP, NULL, NULL);
    bx += bw + 5;

    CreateWindow(L"BUTTON", L"+10",
        WS_CHILD | WS_VISIBLE, bx, by, bw, bh,
        g_hCfgWnd, (HMENU)IDC_CFG_P10, NULL, NULL);

    /* OK / Cancel at bottom */
    by = 255;
    CreateWindow(L"BUTTON", L"OK",
        WS_CHILD | WS_VISIBLE, 60, by, 120, 40,
        g_hCfgWnd, (HMENU)IDC_CFG_OK, NULL, NULL);
    CreateWindow(L"BUTTON", L"Cancel",
        WS_CHILD | WS_VISIBLE, 240, by, 120, 40,
        g_hCfgWnd, (HMENU)IDC_CFG_CANCEL, NULL, NULL);

    /* Enable parent's main window to stay visible but disable input */
    EnableWindow(g_hWnd, FALSE);

    /* Modal message loop for config dialog */
    while (GetMessage(&dmsg, NULL, 0, 0)) {
        TranslateMessage(&dmsg);
        DispatchMessage(&dmsg);
    }

    EnableWindow(g_hWnd, TRUE);
    SetForegroundWindow(g_hWnd);
    return g_cfgResult;
}

/* ============================================================================
 * DoSetupCommon - ensure all registry keys exist for TCP/IP to work.
 * Uses RegCreateKeyEx (not RegOpenKeyEx) so keys are created if missing.
 * Reads the NDIS-generated NetCfgInstanceId GUID from Parms and creates
 * the Interfaces marker key that TCP/IP needs.
 * ============================================================================ */
static void DoSetupCommon(void)
{
    HKEY hKey;
    DWORD dwDisp;
    WCHAR guidBuf[64];
    WCHAR ifPath[128];
    DWORD guidLen, guidType;
    int hasGuid = 0;

    /* 1. Read NDIS-generated NetCfgInstanceId from Parms */
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms",
                     0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        guidLen = sizeof(guidBuf);
        if (RegQueryValueEx(hKey, L"NetCfgInstanceId", NULL, &guidType,
                            (BYTE*)guidBuf, &guidLen) == ERROR_SUCCESS
            && guidType == REG_SZ && guidBuf[0] == L'{') {
            hasGuid = 1;
            {
                WCHAR gm[96];
                wsprintfW(gm, L"GUID: %s", guidBuf);
                AddLog(gm);
            }
        }
        RegCloseKey(hKey);
        AddLog(L"Parms: OK");
    } else {
        AddLog(L"Parms: not found (driver not loaded?)");
    }

    /* 2. Create Comm\Tcpip\Parms\Interfaces\{GUID} marker key.
     *    TCP/IP needs this to associate our adapter with its settings. */
    if (hasGuid) {
        wsprintfW(ifPath, L"Comm\\Tcpip\\Parms\\Interfaces\\%s", guidBuf);
        if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, ifPath,
            0, NULL, 0, 0, NULL, &hKey, &dwDisp) == ERROR_SUCCESS) {
            RegCloseKey(hKey);
            AddLog(L"Interfaces: OK");
        }
    } else {
        AddLog(L"Interfaces: skip (no GUID)");
    }

    /* 3. Ensure Tcpip\Linkage\Bind includes RTL8152B1 */
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\Tcpip\\Linkage",
                       0, NULL, 0, 0, NULL, &hKey, &dwDisp) == ERROR_SUCCESS) {
        BYTE existing[2048];
        DWORD cbData = sizeof(existing), dwType;
        LONG r = RegQueryValueEx(hKey, L"Bind", NULL, &dwType,
                                 existing, &cbData);
        if (r == ERROR_SUCCESS && dwType == REG_MULTI_SZ && cbData > 4) {
            WCHAR *p = (WCHAR*)existing;
            int found = 0;
            while (*p) {
                if (wcscmp(p, L"RTL8152B1") == 0) { found = 1; break; }
                p += wcslen(p) + 1;
            }
            if (!found) {
                WCHAR *end = p;
                wcscpy(end, L"RTL8152B1");
                end += 10; *end = L'\0';
                cbData = (DWORD)((BYTE*)end - existing + sizeof(WCHAR));
                RegSetValueEx(hKey, L"Bind", 0, REG_MULTI_SZ,
                              existing, cbData);
                AddLog(L"Linkage: added");
            } else {
                AddLog(L"Linkage: OK");
            }
        } else {
            WCHAR msz[16];
            wcscpy(msz, L"RTL8152B1");
            msz[9] = L'\0'; msz[10] = L'\0';
            RegSetValueEx(hKey, L"Bind", 0, REG_MULTI_SZ,
                          (const BYTE*)msz, 11 * sizeof(WCHAR));
            AddLog(L"Linkage: created");
        }
        RegCloseKey(hKey);
    }
}

/* ============================================================================
 * DoSetTcpip - write DHCP/static params to RTL8152B1\Parms\Tcpip
 * Keep settings under the hive path already deployed by default.hv.
 * IP fields as REG_MULTI_SZ.  Also set AutoCfg=0 to prevent APIPA.
 * ============================================================================ */
static void DoSetTcpip(int useDhcp, const WCHAR *ip, const WCHAR *mask,
                       const WCHAR *gw)
{
    HKEY hKey;
    DWORD dwDisp;
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\Tcpip",
                     0, NULL, 0, 0, NULL, &hKey, &dwDisp) == ERROR_SUCCESS) {
        RegSetDw(hKey, L"EnableDHCP", useDhcp ? 1 : 0);
        RegSetMsz1(hKey, L"IpAddress",      ip);
        RegSetMsz1(hKey, L"SubnetMask",     mask);
        RegSetMsz1(hKey, L"DefaultGateway", gw);
        RegSetDw(hKey, L"AutoCfg", 0);
        RegSetDw(hKey, L"IPAutoconfigurationEnabled", 0);
        RegSetDw(hKey, L"DontAddDefaultGateway", 0);
        RegSetDw(hKey, L"UseZeroBroadcast", 0);
        RegCloseKey(hKey);
    } else {
        AddLog(L"FAIL: Tcpip key not found");
    }
}

/* ============================================================================
 * DoRebind - ask TCP/IP to re-read registry settings for our adapter.
 *
 * Uses ONLY IOCTL_NDIS_REBIND_ADAPTER (0x17002E).
 * DO NOT use UNBIND (0x170036) - it destroys the TCP/IP binding
 * and cannot be reliably restored, causing GetAdaptersInfo err=232.
 * ============================================================================ */
static void DoRebind(void)
{
    HANDLE hNdis;
    WCHAR msz[32];
    DWORD cbIn, ret;
    BOOL ok;
    WCHAR em[96];

    hNdis = CreateFile(L"NDS0:", GENERIC_READ | GENERIC_WRITE,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_ALWAYS, 0, NULL);
    if (hNdis == INVALID_HANDLE_VALUE) {
        wsprintfW(em, L"NDS0: open FAIL err=%d", GetLastError());
        AddLog(em);
        return;
    }

    /* REBIND only - non-destructive request for TCP/IP to re-read settings */
    memset(msz, 0, sizeof(msz));
    wcscpy(msz, L"RTL8152B1");
    /* msz: "RTL8152B1\0\0" = valid multi_sz with double null */
    cbIn = 11 * sizeof(WCHAR);
    ret = 0;
    ok = DeviceIoControl(hNdis, IOCTL_NDIS_REBIND_ADAPTER,
                         msz, cbIn, NULL, 0, &ret, NULL);
    wsprintfW(em, L"Rebind: %s err=%d", ok ? L"OK" : L"FAIL",
              GetLastError());
    AddLog(em);

    CloseHandle(hNdis);
}

/* ============================================================================
 * PingAddr - ping single IP via IcmpSendEcho (loaded at runtime)
 * Returns TRUE on success, sets *pRTT to round-trip time in ms
 * ============================================================================ */
static BOOL PingAddr(DWORD ipNBO, DWORD *pRTT)
{
    static HMODULE hMod = NULL;
    static PFN_IcmpCreate pfnCreate = NULL;
    static PFN_IcmpSend   pfnSend   = NULL;
    static PFN_IcmpClose  pfnClose  = NULL;
    static int bInited = 0;
    HANDLE hIcmp;
    char sendBuf[32];
    BYTE replyBuf[512];
    DWORD ret;

    *pRTT = 0;

    if (!bInited) {
        bInited = 1;
        hMod = LoadLibrary(L"iphlpapi.dll");
        if (hMod) {
            pfnCreate = (PFN_IcmpCreate)GetProcAddress(hMod, L"IcmpCreateFile");
            pfnSend   = (PFN_IcmpSend)  GetProcAddress(hMod, L"IcmpSendEcho");
            pfnClose  = (PFN_IcmpClose) GetProcAddress(hMod, L"IcmpCloseHandle");
        }
    }

    if (!pfnCreate || !pfnSend || !pfnClose) {
        AddLog(L"ICMP API not available");
        return FALSE;
    }

    hIcmp = pfnCreate();
    if (hIcmp == INVALID_HANDLE_VALUE) {
        AddLog(L"IcmpCreateFile failed");
        return FALSE;
    }

    memset(sendBuf, 0x41, sizeof(sendBuf));
    ret = pfnSend(hIcmp, ipNBO, sendBuf, 32, NULL,
                  replyBuf, sizeof(replyBuf), 3000);
    pfnClose(hIcmp);

    if (ret > 0) {
        /* replyBuf layout: [Address 4B][Status 4B][RTT 4B]... */
        DWORD status = *(DWORD*)(replyBuf + 4);
        *pRTT = *(DWORD*)(replyBuf + 8);
        return (status == 0);
    }
    return FALSE;
}

/* ============================================================================
 * PingSrvThread - ICMP echo server (ping responder)
 * Listens on raw ICMP socket, replies to echo requests.
 * If raw sockets unavailable, falls back to UDP echo on port 7.
 * ============================================================================ */
static DWORD WINAPI PingSrvThread(LPVOID param)
{
    SOCKET s;
    struct sockaddr_in bindAddr, fromAddr;
    int fromLen, recvLen;
    BYTE buf[1500];
    DWORD rcvTimeo;
    int pingReplied = 0;
    int useUdp = 0;
    WCHAR em[128];
    (void)param;

    /* Try raw ICMP first */
    s = socket(AF_INET, SOCK_RAW, IPPROTO_ICMP);
    if (s == INVALID_SOCKET) {
        int rawErr = WSAGetLastError();
        wsprintfW(em, L"PingSrv: raw FAIL err=%d, trying UDP:7", rawErr);
        AddLog(em);

        /* Fallback: UDP echo on port 7 */
        s = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP);
        if (s == INVALID_SOCKET) {
            wsprintfW(em, L"PingSrv: UDP socket FAIL err=%d", WSAGetLastError());
            AddLog(em);
            g_bPingSrvRun = 0;
            return 1;
        }
        useUdp = 1;
    }
    g_sPingSrv = s;

    memset(&bindAddr, 0, sizeof(bindAddr));
    bindAddr.sin_family = AF_INET;
    bindAddr.sin_addr.s_addr = 0;
    if (useUdp) bindAddr.sin_port = htons(7);

    if (bind(s, (struct sockaddr*)&bindAddr, sizeof(bindAddr)) != 0) {
        wsprintfW(em, L"PingSrv: bind FAIL err=%d", WSAGetLastError());
        AddLog(em);
        closesocket(s);
        g_sPingSrv = INVALID_SOCKET;
        g_bPingSrvRun = 0;
        return 1;
    }

    rcvTimeo = 1000;
    setsockopt(s, SOL_SOCKET, SO_RCVTIMEO,
               (const char*)&rcvTimeo, sizeof(rcvTimeo));

    if (useUdp)
        AddLog(L"PingSrv: UDP echo on port 7 (use: nping --udp -p 7 IP)");
    else
        AddLog(L"PingSrv: listening for ICMP echo...");

    while (g_bPingSrvRun) {
        fromLen = sizeof(fromAddr);
        recvLen = recvfrom(s, (char*)buf, sizeof(buf), 0,
                           (struct sockaddr*)&fromAddr, &fromLen);
        if (recvLen == SOCKET_ERROR) {
            int se = WSAGetLastError();
            if (se == WSAETIMEDOUT || se == WSAEWOULDBLOCK) continue;
            if (!g_bPingSrvRun) break;
            wsprintfW(em, L"PingSrv: recvfrom err=%d", se);
            AddLog(em);
            continue;
        }

        if (useUdp) {
            /* UDP echo: send back exactly what we received */
            DWORD srcIp;
            sendto(s, (char*)buf, recvLen, 0,
                   (struct sockaddr*)&fromAddr, sizeof(fromAddr));
            pingReplied++;
            srcIp = fromAddr.sin_addr.s_addr;
            if (pingReplied <= 5 || (pingReplied % 100 == 0)) {
                wsprintfW(em, L"UdpEcho: #%d from %d.%d.%d.%d (%dB)",
                          pingReplied,
                          srcIp & 0xFF, (srcIp >> 8) & 0xFF,
                          (srcIp >> 16) & 0xFF, (srcIp >> 24) & 0xFF,
                          recvLen);
                AddLog(em);
            }
        } else {
            /* Raw ICMP: parse and reply */
            int ihl;
            AX_ICMP_HDR *icmp;
            int icmpLen;

            if ((buf[0] >> 4) == 4)
                ihl = (buf[0] & 0x0F) * 4;
            else
                ihl = 0;

            if (recvLen < ihl + 8) continue;

            icmp = (AX_ICMP_HDR *)(buf + ihl);
            icmpLen = recvLen - ihl;

            if (icmp->type == AX_ICMP_ECHO_REQ && icmp->code == 0) {
                BYTE reply[1500];
                DWORD srcIp;

                if (icmpLen > (int)sizeof(reply)) icmpLen = (int)sizeof(reply);

                ((AX_ICMP_HDR*)reply)->type = AX_ICMP_ECHO_REP;
                ((AX_ICMP_HDR*)reply)->cksum = 0;
                ((AX_ICMP_HDR*)reply)->cksum = IcmpCksum(reply, icmpLen);

                sendto(s, (char*)reply, icmpLen, 0,
                       (struct sockaddr*)&fromAddr, sizeof(fromAddr));

                pingReplied++;
                srcIp = fromAddr.sin_addr.s_addr;

                if (pingReplied <= 5 || (pingReplied % 100 == 0)) {
                    wsprintfW(em, L"PingSrv: #%d from %d.%d.%d.%d (%dB)",
                              pingReplied,
                              srcIp & 0xFF, (srcIp >> 8) & 0xFF,
                              (srcIp >> 16) & 0xFF, (srcIp >> 24) & 0xFF,
                              icmpLen);
                    AddLog(em);
                }
            }
        }
    }

    closesocket(s);
    g_sPingSrv = INVALID_SOCKET;
    g_bPingSrvRun = 0;
    wsprintfW(em, L"PingSrv: stopped (%d replies)", pingReplied);
    AddLog(em);
    return 0;
}

/* ============================================================================
 * DoShowStatus - show all adapters: IP, MAC, gateway, DHCP mode
 * ============================================================================ */
static void DoShowStatus(void)
{
    ULONG bufLen = 16384;
    PIP_ADAPTER_INFO pAI, pCur;
    DWORD dwRet;
    int count = 0;
    HKEY hKey;
    DWORD regDhcp, cb, dwType;

    AddLog(L"=== Network Status ===");

    /* DHCP mode from registry */
    regDhcp = 99; cb = sizeof(DWORD);
    if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\Tcpip",
                     0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        RegQueryValueEx(hKey, L"EnableDHCP", NULL, &dwType,
                        (BYTE*)&regDhcp, &cb);
        RegCloseKey(hKey);
    }
    AddLog(regDhcp == 1 ? L"Config: DHCP" :
           regDhcp == 0 ? L"Config: Static" : L"Config: (not set)");

    pAI = (PIP_ADAPTER_INFO)LocalAlloc(LPTR, bufLen);
    if (!pAI) { AddLog(L"Memory error"); return; }

    dwRet = GetAdaptersInfo(pAI, &bufLen);
    if (dwRet != ERROR_SUCCESS) {
        WCHAR em[80];
        wsprintfW(em, L"GetAdaptersInfo err=%d", dwRet);
        AddLog(em);
        LocalFree(pAI);
        return;
    }

    for (pCur = pAI; pCur; pCur = pCur->Next) {
        WCHAR info[256], wN[128], wIP[32], wMask[32], wGW[32];

        AtoW(pCur->AdapterName, wN, 128);
        AtoW(pCur->IpAddressList.IpAddress.String, wIP, 32);
        AtoW(pCur->IpAddressList.IpMask.String, wMask, 32);
        AtoW(pCur->GatewayList.IpAddress.String, wGW, 32);

        if (pCur->AddressLength == 6) {
            wsprintfW(info, L"[%s] MAC %02X:%02X:%02X:%02X:%02X:%02X",
                      wN, pCur->Address[0], pCur->Address[1],
                      pCur->Address[2], pCur->Address[3],
                      pCur->Address[4], pCur->Address[5]);
            AddLog(info);

            if (pCur->Address[0] == 0x02 && pCur->Address[1] == 0x0B &&
                pCur->Address[2] == 0x95 && pCur->Address[3] == 0x17 &&
                pCur->Address[4] == 0x90 && pCur->Address[5] == 0x01) {
                AddLog(L"  WARNING: fallback MAC in use; USBDeviceAttach did not provide real hardware state");
                AddLog(L"  ACTION: run Arm, then do a real USB unplug/replug, then check Driver Log again");
            }
        } else {
            wsprintfW(info, L"[%s]", wN);
            AddLog(info);
        }

        wsprintfW(info, L"  IP=%s  Mask=%s", wIP, wMask);
        AddLog(info);
        wsprintfW(info, L"  GW=%s  DHCP=%s", wGW,
                  pCur->DhcpEnabled ? L"yes" : L"no");
        AddLog(info);

        if (pCur->DhcpEnabled) {
            WCHAR wDS[32];
            AtoW(pCur->DhcpServer.IpAddress.String, wDS, 32);
            wsprintfW(info, L"  DHCP srv=%s", wDS);
            AddLog(info);
        }

        count++;
    }

    if (count == 0) AddLog(L"No adapters found");
    AddLog(L"======================");
}

static BOOL FindRtl8152Adapter(DWORD *ifIdx, WCHAR *ipBuf, DWORD cchIpBuf,
                               BOOL *dhcpEnabled)
{
    ULONG bufLen = 16384;
    PIP_ADAPTER_INFO pAI, pCur;
    BOOL found;

    if (ifIdx) *ifIdx = 0;
    if (ipBuf && cchIpBuf) ipBuf[0] = L'\0';
    if (dhcpEnabled) *dhcpEnabled = FALSE;

    pAI = (PIP_ADAPTER_INFO)LocalAlloc(LPTR, bufLen);
    if (!pAI) {
        return FALSE;
    }

    found = FALSE;
    if (GetAdaptersInfo(pAI, &bufLen) == ERROR_SUCCESS) {
        for (pCur = pAI; pCur; pCur = pCur->Next) {
            WCHAR wN[128];
            AtoW(pCur->AdapterName, wN, 128);
            if (wcscmp(wN, L"RTL8152B1") == 0) {
                if (ifIdx) *ifIdx = pCur->Index;
                if (ipBuf && cchIpBuf) {
                    AtoW(pCur->IpAddressList.IpAddress.String, ipBuf, (int)cchIpBuf);
                    ipBuf[cchIpBuf - 1] = L'\0';
                }
                if (dhcpEnabled) *dhcpEnabled = pCur->DhcpEnabled ? TRUE : FALSE;
                found = TRUE;
                break;
            }
        }
    }

    LocalFree(pAI);
    return found;
}

static BOOL IsZeroIpStringW(const WCHAR *ip)
{
    if (!ip || !ip[0]) {
        return TRUE;
    }
    return (wcscmp(ip, L"0.0.0.0") == 0);
}

/* ============================================================================
 * FindLatestLog - find highest-numbered rtl8152_N.log on MediaCard
 * ============================================================================ */
static int FindLatestLog(void)
{
    int n, latest = 0;
    WCHAR tmp[64];
    for (n = 1; n < 9999; n++) {
        wsprintfW(tmp, L"\\MediaCard\\rtl8152_%d.log", n);
        if (GetFileAttributes(tmp) == 0xFFFFFFFF) break;
        latest = n;
    }
    return latest;
}

static void AnalyzeDriverLogBuffer(const char *buf, DWORD rd)
{
    int hasInitDone = 0;
    int hasEp0Err = 0;
    int hasNullXfer = 0;
    int hasPlaCr = 0;
    int hasLinkUp = 0;
    int hasLinkMonUp = 0;

    if (!buf || rd == 0)
        return;

    hasInitDone = (strstr(buf, "INIT: RtlInitChip DONE") != NULL);
    hasEp0Err = (strstr(buf, "USB_WR: ERR=") != NULL) ||
                (strstr(buf, "USB_RD: ERR=") != NULL);
    hasNullXfer = (strstr(buf, "hXfer=NULL") != NULL);
    hasPlaCr = (strstr(buf, "INIT: PLA_CR=") != NULL);
    hasLinkUp = (strstr(buf, "link=UP") != NULL);
    hasLinkMonUp = (strstr(buf, "LinkMon: link UP") != NULL);

    AddLog(L"--- Driver Log Analysis ---");

    if (hasEp0Err || hasNullXfer) {
        AddLog(L"VERDICT: FAIL - USB control register access is still broken");
        if (hasEp0Err)
            AddLog(L"Detail: USB_RD/USB_WR ERR seen in log");
        if (hasNullXfer)
            AddLog(L"Detail: IssueVendorTransfer returned NULL");
        return;
    }

    if (!hasInitDone) {
        AddLog(L"VERDICT: INCOMPLETE - RtlInitChip did not finish");
        return;
    }

    if (hasPlaCr)
        AddLog(L"Check: PLA_CR readback present");
    else
        AddLog(L"Check: PLA_CR readback missing");

    if (hasLinkMonUp || hasLinkUp)
        AddLog(L"VERDICT: INIT OK - link up observed in driver log");
    else
        AddLog(L"VERDICT: INIT OK - but link up not yet observed");
}

/* ============================================================================
 * DoShowDriverLog - read and display latest rtl8152_N.log
 * ============================================================================ */
static void DoShowDriverLog(void)
{
    HANDLE hFile;
    DWORD sz, rd;
    char *buf;
    int start, i, latest;
    WCHAR logPath[64];

    AddLog(L"--- Driver Log ---");
    latest = FindLatestLog();
    if (latest == 0) {
        AddLog(L"No driver log found");
        AddLog(L"This usually means USBDeviceAttach has not run yet");
        AddLog(L"Use Arm, then do a real USB unplug/replug, then retry Log");
        return;
    }

    wsprintfW(logPath, L"\\MediaCard\\rtl8152_%d.log", latest);
    {
        WCHAR lm[80];
        wsprintfW(lm, L"Log #%d", latest);
        AddLog(lm);
    }

    hFile = CreateFile(logPath, GENERIC_READ,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_EXISTING, 0, NULL);
    if (hFile == INVALID_HANDLE_VALUE) {
        AddLog(L"Cannot open driver log");
        return;
    }

    sz = GetFileSize(hFile, NULL);
    if (sz == 0 || sz > 65536) {
        WCHAR em[80];
        wsprintfW(em, L"Log size: %d bytes (skipped)", sz);
        AddLog(em);
        CloseHandle(hFile);
        return;
    }

    buf = (char*)LocalAlloc(LPTR, sz + 1);
    if (!buf) { CloseHandle(hFile); return; }
    ReadFile(hFile, buf, sz, &rd, NULL);
    buf[rd] = '\0';
    CloseHandle(hFile);

    start = 0;
    for (i = 0; i <= (int)rd; i++) {
        if (buf[i] == '\n' || buf[i] == '\0') {
            if (i > start) {
                WCHAR wline[512];
                int j, len = i - start;
                if (len > 0 && buf[start + len - 1] == '\r') len--;
                if (len > 500) len = 500;
                for (j = 0; j < len; j++)
                    wline[j] = (WCHAR)(unsigned char)buf[start + j];
                wline[len] = 0;
                if (g_hList) {
                    LRESULT idx = SendMessage(g_hList, LB_ADDSTRING,
                                              0, (LPARAM)wline);
                    SendMessage(g_hList, LB_SETTOPINDEX, idx, 0);
                }
            }
            start = i + 1;
        }
    }
    AnalyzeDriverLogBuffer(buf, rd);
    UpdateWindow(g_hList);
    LocalFree(buf);
    AddLog(L"--- End ---");
}

/* ============================================================================
 * DoSaveListbox - save listbox to \MediaCard\netconfig_save.txt
 * ============================================================================ */
static void DoSaveListbox(void)
{
    HANDLE hFile;
    DWORD written;
    LRESULT count, i;
    WCHAR line[512];
    char abuf[1024];
    int j;

    if (!g_hList) return;
    count = SendMessage(g_hList, LB_GETCOUNT, 0, 0);
    if (count <= 0) { AddLog(L"Nothing to save"); return; }

    hFile = CreateFile(L"\\MediaCard\\netconfig_save.txt",
                       GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
                       FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) {
        AddLog(L"Cannot create save file");
        return;
    }

    for (i = 0; i < count; i++) {
        LRESULT len = SendMessage(g_hList, LB_GETTEXT, (WPARAM)i, (LPARAM)line);
        if (len > 0 && len < 500) {
            j = 0;
            while (line[j] && j < 1020) { abuf[j] = (char)line[j]; j++; }
            abuf[j] = '\r'; abuf[j + 1] = '\n';
            WriteFile(hFile, abuf, j + 2, &written, NULL);
        }
    }
    CloseHandle(hFile);
    {
        WCHAR sm[80];
        wsprintfW(sm, L"Saved %d lines", (int)count);
        AddLog(sm);
    }
}

/* ============================================================================
 * DHCPThread - configure DHCP + rebind + wait + show status
 * ============================================================================ */
static DWORD WINAPI DHCPThread(LPVOID param)
{
    int latest;
    (void)param;

    AddLog(L"====== DHCP Setup ======");

    latest = FindLatestLog();
    if (latest > 0) {
        WCHAR sm[64];
        wsprintfW(sm, L"Driver log #%d present", latest);
        AddLog(sm);
    } else {
        AddLog(L"WARNING: no driver log found");
    }

    DoSetupCommon();
    DoSetTcpip(1, L"0.0.0.0", L"0.0.0.0", L"0.0.0.0");
    AddLog(L"DHCP mode set");

    DoRebind();

    AddLog(L"Waiting 5s for DHCP...");
    Sleep(5000);

    DoShowStatus();

    AddLog(L"Done!");
    InterlockedExchange(&g_bBusy, 0);
    return 0;
}

/* ============================================================================
 * StaticApplyThread - apply static config from g_cfg[], rebind, show status
 * ============================================================================ */
static DWORD WINAPI StaticApplyThread(LPVOID param)
{
    WCHAR ip[32], mask[32], gw[32], dns[32], m[128];
    HKEY hKey;
    (void)param;

    AddLog(L"====== Static IP Setup ======");

    CfgToStr(0, ip);   wsprintfW(m, L"IP:   %s", ip);   AddLog(m);
    CfgToStr(1, mask);  wsprintfW(m, L"Mask: %s", mask); AddLog(m);
    CfgToStr(2, gw);   wsprintfW(m, L"GW:   %s", gw);   AddLog(m);
    CfgToStr(3, dns);  wsprintfW(m, L"DNS:  %s", dns);  AddLog(m);

    DoSetupCommon();
    DoSetTcpip(0, ip, mask, gw);

    /* Write DNS1 - use RegCreateKeyEx in case DoSetTcpip path differs */
    {
        DWORD dwDisp2;
        if (RegCreateKeyEx(HKEY_LOCAL_MACHINE,
                           L"Comm\\RTL8152B1\\Parms\\Tcpip",
                           0, NULL, 0, 0, NULL, &hKey, &dwDisp2) == ERROR_SUCCESS) {
            RegSetMsz1(hKey, L"DNS", dns);
            RegCloseKey(hKey);
        }
    }

    AddLog(L"Static mode set");
    DoRebind();

    /* Direct IP assignment via AddIPAddress as backup.
     * Even if registry-based rebind doesn't take effect immediately,
     * this forces the IP onto the adapter right now.
     * Try twice with a delay - adapter may need time after rebind. */
    {
        ULONG bufLen = 16384;
        PIP_ADAPTER_INFO pAI, pCur;
        DWORD ifIdx = 0;
        int attempt;

        for (attempt = 0; attempt < 2 && ifIdx == 0; attempt++) {
            if (attempt > 0) {
                AddLog(L"Retry GetAdaptersInfo...");
                Sleep(2000);
            }
            bufLen = 16384;
            pAI = (PIP_ADAPTER_INFO)LocalAlloc(LPTR, bufLen);
            if (pAI && GetAdaptersInfo(pAI, &bufLen) == ERROR_SUCCESS) {
                for (pCur = pAI; pCur; pCur = pCur->Next) {
                    WCHAR wN[128];
                    AtoW(pCur->AdapterName, wN, 128);
                    if (wcscmp(wN, L"RTL8152B1") == 0) {
                        ifIdx = pCur->Index;
                        break;
                    }
                }
            }
            if (pAI) LocalFree(pAI);
        }

        if (ifIdx != 0) {
            ULONG nteCtx = 0, nteInst = 0;
            DWORD ipAddr, ipMask;
            WCHAR em2[96];
            ipAddr = ParseIPA2(ip);
            ipMask = ParseIPA2(mask);
            {
                DWORD r = AddIPAddress(ipAddr, ipMask, ifIdx,
                                       &nteCtx, &nteInst);
                wsprintfW(em2, L"AddIPAddress: rc=%d ifIdx=%d", r, ifIdx);
                AddLog(em2);
            }
        } else {
            AddLog(L"AddIPAddress: adapter not found (reboot?)");
        }
    }

    AddLog(L"Waiting 3s...");
    Sleep(3000);

    DoShowStatus();

    AddLog(L"Done!");
    InterlockedExchange(&g_bBusy, 0);
    return 0;
}

/* ============================================================================
 * PingThread - ping gateway + 8.8.8.8
 * ============================================================================ */
static DWORD WINAPI PingThread(LPVOID param)
{
    ULONG bufLen = 16384;
    PIP_ADAPTER_INFO pAI, pCur;
    DWORD gwAddr = 0, rtt;
    WCHAR ipStr[32], msg[128];
    (void)param;

    AddLog(L"=== Ping Test ===");

    /* Find gateway from adapter info */
    pAI = (PIP_ADAPTER_INFO)LocalAlloc(LPTR, bufLen);
    if (pAI && GetAdaptersInfo(pAI, &bufLen) == ERROR_SUCCESS) {
        for (pCur = pAI; pCur; pCur = pCur->Next) {
            DWORD g = ParseIPA(pCur->GatewayList.IpAddress.String);
            if (g != 0) { gwAddr = g; break; }
        }
    }
    if (pAI) LocalFree(pAI);

    /* Ping gateway */
    if (gwAddr != 0) {
        FormatIP(gwAddr, ipStr);
        wsprintfW(msg, L"Ping GW %s ...", ipStr);
        AddLog(msg);
        if (PingAddr(gwAddr, &rtt)) {
            wsprintfW(msg, L"  Reply: %d ms", rtt);
        } else {
            wsprintfW(msg, L"  No reply (timeout 3s)");
        }
        AddLog(msg);
    } else {
        AddLog(L"No gateway - trying 192.168.1.1");
        gwAddr = ParseIPA("192.168.1.1");
        if (PingAddr(gwAddr, &rtt)) {
            wsprintfW(msg, L"  192.168.1.1 reply: %d ms", rtt);
        } else {
            wsprintfW(msg, L"  192.168.1.1 no reply");
        }
        AddLog(msg);
    }

    /* Ping 8.8.8.8 (internet check) */
    {
        DWORD extIP = ParseIPA("8.8.8.8");
        AddLog(L"Ping 8.8.8.8 ...");
        if (PingAddr(extIP, &rtt)) {
            wsprintfW(msg, L"  Reply: %d ms (internet OK)", rtt);
        } else {
            wsprintfW(msg, L"  No reply (no internet)");
        }
        AddLog(msg);
    }

    AddLog(L"=== Ping Done ===");
    InterlockedExchange(&g_bBusy, 0);
    return 0;
}

static DWORD WINAPI AutoStartupThread(LPVOID param)
{
    DWORD mode;
    DWORD ifIdx;
    WCHAR ip[32];
    BOOL dhcpEnabled;
    DWORD waitLoops;

    mode = (DWORD)param;
    waitLoops = 0;

    if (mode == NETCFG_STARTUP_MODE_STATIC) {
        AddLog(L"AutoStart: waiting for RTL8152B1 to apply static config");
    } else if (mode == NETCFG_STARTUP_MODE_DHCP) {
        AddLog(L"AutoStart: waiting for RTL8152B1 to apply DHCP config");
    } else {
        return 0;
    }

    for (;;) {
        if (FindRtl8152Adapter(&ifIdx, ip, sizeof(ip) / sizeof(ip[0]), &dhcpEnabled)) {
            if (mode == NETCFG_STARTUP_MODE_STATIC) {
                if (!IsZeroIpStringW(ip) && !dhcpEnabled) {
                    AddLog(L"AutoStart: RTL8152B1 already has static IP, no action needed");
                    break;
                }

                AddLog(L"AutoStart: RTL8152B1 detected, applying static config");
                CfgDefaults();
                if (InterlockedExchange(&g_bBusy, 1) == 0) {
                    StaticApplyThread(NULL);
                    break;
                }
            } else if (mode == NETCFG_STARTUP_MODE_DHCP) {
                if (dhcpEnabled && !IsZeroIpStringW(ip)) {
                    AddLog(L"AutoStart: RTL8152B1 already has DHCP address, no action needed");
                    break;
                }

                AddLog(L"AutoStart: RTL8152B1 detected, applying DHCP config");
                if (InterlockedExchange(&g_bBusy, 1) == 0) {
                    DHCPThread(NULL);
                    break;
                }
            }
        }

        waitLoops++;
        if ((waitLoops % 10) == 0) {
            AddLog(L"AutoStart: waiting for adapter...");
        }
        Sleep(1000);
    }

    return 0;
}

static DWORD WINAPI NetCheckThread(LPVOID param)
{
    DWORD rtt;
    DWORD statusCode;
    DWORD elapsedMs;
    DWORD bodyBytes;
    WCHAR msg[160];
    (void)param;

    AddLog(L"=== Net Check ===");
    DoShowStatus();

    AddLog(L"DNS:");
    LogResolvedHost("connectivitycheck.gstatic.com");
    LogResolvedHost("neverssl.com");

    AddLog(L"HTTP:");
    if (HttpSimpleGet("connectivitycheck.gstatic.com", "/generate_204",
                      &statusCode, &elapsedMs, &bodyBytes)) {
        wsprintfW(msg, L"  connectivitycheck.gstatic.com status=%u time=%u ms body=%uB",
                  statusCode, elapsedMs, bodyBytes);
        AddLog(msg);
    } else if (HttpSimpleGet("neverssl.com", "/",
                             &statusCode, &elapsedMs, &bodyBytes)) {
        wsprintfW(msg, L"  neverssl.com status=%u time=%u ms body=%uB",
                  statusCode, elapsedMs, bodyBytes);
        AddLog(msg);
    } else {
        AddLog(L"  HTTP FAIL");
    }

    AddLog(L"ICMP:");
    if (PingAddr(ParseIPA("8.8.8.8"), &rtt))
        wsprintfW(msg, L"  8.8.8.8 reply=%u ms", rtt);
    else
        wsprintfW(msg, L"  8.8.8.8 no reply");
    AddLog(msg);

    AddLog(L"=== Net Check Done ===");
    InterlockedExchange(&g_bBusy, 0);
    return 0;
}

static DWORD WINAPI DnsTestThread(LPVOID param)
{
    (void)param;
    AddLog(L"=== DNS Test ===");
    LogResolvedHost("example.com");
    LogResolvedHost("neverssl.com");
    LogResolvedHost("connectivitycheck.gstatic.com");
    AddLog(L"=== DNS Done ===");
    InterlockedExchange(&g_bBusy, 0);
    return 0;
}

static DWORD WINAPI HttpTestThread(LPVOID param)
{
    DWORD statusCode;
    DWORD elapsedMs;
    DWORD bodyBytes;
    WCHAR msg[160];
    (void)param;

    AddLog(L"=== HTTP Test ===");
    if (HttpSimpleGet("connectivitycheck.gstatic.com", "/generate_204",
                      &statusCode, &elapsedMs, &bodyBytes)) {
        wsprintfW(msg, L"  connectivitycheck.gstatic.com status=%u time=%u ms body=%uB",
                  statusCode, elapsedMs, bodyBytes);
        AddLog(msg);
    } else {
        AddLog(L"  connectivitycheck.gstatic.com FAIL");
    }

    if (HttpSimpleGet("neverssl.com", "/",
                      &statusCode, &elapsedMs, &bodyBytes)) {
        wsprintfW(msg, L"  neverssl.com status=%u time=%u ms body=%uB",
                  statusCode, elapsedMs, bodyBytes);
        AddLog(msg);
    } else {
        AddLog(L"  neverssl.com FAIL");
    }

    AddLog(L"=== HTTP Done ===");
    InterlockedExchange(&g_bBusy, 0);
    return 0;
}

static DWORD WINAPI Download10sThread(LPVOID param)
{
    static const char *hosts[] = {
        "speedtest.tele2.net",
        "ipv4.download.thinkbroadband.com"
    };
    static const char *paths[] = {
        "/10MB.zip",
        "/5MB.zip"
    };
    DWORD elapsedMs;
    DWORD bodyBytes;
    DWORD statusCode;
    WCHAR msg[192];
    int i;
    int ok;

    (void)param;
    AddLog(L"=== DL10s Test ===");
    ok = 0;

    for (i = 0; i < 2; i++) {
        WCHAR whost[96];
        AtoW(hosts[i], whost, 96);
        wsprintfW(msg, L"  Trying %s", whost);
        AddLog(msg);

        if (HttpDownloadForMs(hosts[i], paths[i], 10000,
                              &elapsedMs, &bodyBytes, &statusCode)) {
            DWORD kbps1000;
            DWORD mbps100;

            if (elapsedMs == 0) elapsedMs = 1;
            kbps1000 = (bodyBytes * 1000UL) / elapsedMs;
            mbps100 = (kbps1000 * 8UL) / 10000UL;
            wsprintfW(msg,
                      L"  status=%u bytes=%u time=%u ms speed=%u.%02u Mbps",
                      statusCode, bodyBytes, elapsedMs,
                      mbps100 / 100, mbps100 % 100);
            AddLog(msg);
            ok = 1;
            break;
        }

        wsprintfW(msg, L"  %s FAIL", whost);
        AddLog(msg);
    }

    if (!ok)
        AddLog(L"  Download mirrors unavailable");

    AddLog(L"=== DL10s Done ===");
    InterlockedExchange(&g_bBusy, 0);
    return 0;
}

/* ============================================================================
 * WndProc
 * ============================================================================ */
static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg,
                                WPARAM wParam, LPARAM lParam)
{
    switch (msg) {
    case WM_CREATE: {
        RECT rc;
        int w, h, bw, bh, by, gap, x;

        /* Create all controls at (0,0,0,0) - WM_SIZE will layout them */
        g_hList = CreateWindow(L"LISTBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOINTEGRALHEIGHT,
            0, 0, 10, 10, hWnd, (HMENU)IDC_LIST, NULL, NULL);

        CreateWindow(L"BUTTON", L"DHCP",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_DHCP, NULL, NULL);
        CreateWindow(L"BUTTON", L"Static",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_STATIC, NULL, NULL);
        CreateWindow(L"BUTTON", L"Status",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_STATUS, NULL, NULL);
        CreateWindow(L"BUTTON", L"Ping",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_PING, NULL, NULL);
        CreateWindow(L"BUTTON", L"Log",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_LOG, NULL, NULL);
        CreateWindow(L"BUTTON", L"Clr",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_CLEAR, NULL, NULL);
        CreateWindow(L"BUTTON", L"Save",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_SAVE, NULL, NULL);
        CreateWindow(L"BUTTON", L"Net",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_PINGSRV, NULL, NULL);
        CreateWindow(L"BUTTON", L"DNS",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_DNS, NULL, NULL);
        CreateWindow(L"BUTTON", L"HTTP",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_HTTP, NULL, NULL);
        CreateWindow(L"BUTTON", L"DL10s",
            WS_CHILD | WS_VISIBLE, 0,0,10,10, hWnd,
            (HMENU)IDC_BTN_DL10S, NULL, NULL);
        /* Immediate layout from actual client rect */
        GetClientRect(hWnd, &rc);
        w = rc.right; h = rc.bottom;
        bh = 28; gap = 3; by = h - bh - 4;
        bw = (w - 10 - 10 * gap) / 11;
        if (bw < 28) bw = 28;
        MoveWindow(g_hList, 4, 4, w - 8, by - 8, TRUE);
        x = 4;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_DHCP),    x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_STATIC),  x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_STATUS),  x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_PING),    x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_LOG),     x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_CLEAR),   x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_SAVE),    x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_PINGSRV), x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_DNS),     x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_HTTP),    x, by, bw, bh, TRUE); x += bw + gap;
        MoveWindow(GetDlgItem(hWnd, IDC_BTN_DL10S),   x, by, bw, bh, TRUE);

        AddLog(L"Ready. AutoStart watcher is active; manual DHCP/Static still available.");
        break;
    }

    case WM_SIZE: {
        int w = LOWORD(lParam);
        int h = HIWORD(lParam);
        int bw, gap, bh, by, x, i;
        UINT ids[11];
        if (w < 20 || h < 20) break;
        ids[0] = IDC_BTN_DHCP;   ids[1] = IDC_BTN_STATIC;
        ids[2] = IDC_BTN_STATUS; ids[3] = IDC_BTN_PING;
        ids[4] = IDC_BTN_LOG;    ids[5] = IDC_BTN_CLEAR;
        ids[6] = IDC_BTN_SAVE;   ids[7] = IDC_BTN_PINGSRV;
        ids[8] = IDC_BTN_DNS;    ids[9] = IDC_BTN_HTTP;
        ids[10] = IDC_BTN_DL10S;
        bh = 28; gap = 3; by = h - bh - 4;
        bw = (w - 10 - 10 * gap) / 11;
        if (bw < 28) bw = 28;
        MoveWindow(g_hList, 4, 4, w - 8, by - 8, TRUE);
        x = 4;
        for (i = 0; i < 11; i++) {
            MoveWindow(GetDlgItem(hWnd, ids[i]), x, by, bw, bh, TRUE);
            x += bw + gap;
        }
        break;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_BTN_DHCP:
            if (InterlockedExchange(&g_bBusy, 1) == 0)
                CreateThread(NULL, 0, DHCPThread, NULL, 0, NULL);
            else
                AddLog(L"Busy...");
            break;
        case IDC_BTN_STATIC:
            if (g_bBusy) {
                AddLog(L"Busy...");
            } else {
                HINSTANCE hI = (HINSTANCE)GetWindowLong(hWnd, GWL_HINSTANCE);
                if (ShowStaticConfigDialog(hI)) {
                    InterlockedExchange(&g_bBusy, 1);
                    CreateThread(NULL, 0, StaticApplyThread, NULL, 0, NULL);
                } else {
                    AddLog(L"Static config cancelled");
                }
            }
            break;
        case IDC_BTN_STATUS:
            DoShowStatus();
            break;
        case IDC_BTN_PING:
            if (InterlockedExchange(&g_bBusy, 1) == 0)
                CreateThread(NULL, 0, PingThread, NULL, 0, NULL);
            else
                AddLog(L"Busy...");
            break;
        case IDC_BTN_LOG:
            DoShowDriverLog();
            break;
        case IDC_BTN_CLEAR:
            SendMessage(g_hList, LB_RESETCONTENT, 0, 0);
            break;
        case IDC_BTN_SAVE:
            DoSaveListbox();
            break;
        case IDC_BTN_PINGSRV:
            if (InterlockedExchange(&g_bBusy, 1) == 0)
                CreateThread(NULL, 0, NetCheckThread, NULL, 0, NULL);
            else
                AddLog(L"Busy...");
            break;
        case IDC_BTN_DNS:
            if (InterlockedExchange(&g_bBusy, 1) == 0)
                CreateThread(NULL, 0, DnsTestThread, NULL, 0, NULL);
            else
                AddLog(L"Busy...");
            break;
        case IDC_BTN_HTTP:
            if (InterlockedExchange(&g_bBusy, 1) == 0)
                CreateThread(NULL, 0, HttpTestThread, NULL, 0, NULL);
            else
                AddLog(L"Busy...");
            break;
        case IDC_BTN_DL10S:
            if (InterlockedExchange(&g_bBusy, 1) == 0)
                CreateThread(NULL, 0, Download10sThread, NULL, 0, NULL);
            else
                AddLog(L"Busy...");
            break;
        case IDC_BTN_PROBE:
            DoProbeDriver();
            break;
        case IDC_BTN_STAGE:
            DoStageUsbDebug();
            break;
        case IDC_BTN_KICK:
            DoKickNdis();
            break;
        case IDC_BTN_TESTDLL:
            DoTestDllLoad();
            break;
        case IDC_BTN_VBUS:
            DoEhciTakeover();
            break;
        }
        break;

    case WM_DESTROY:
        if (g_bPingSrvRun) {
            g_bPingSrvRun = 0;
            if (g_sPingSrv != INVALID_SOCKET)
                closesocket(g_sPingSrv);
        }
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

/* ============================================================================
 * WinMain
 * ============================================================================ */
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrev, LPTSTR lpCmd, int nShow)
{
    WNDCLASS wc;
    MSG msg;
    int autoStarted;

    (void)hPrev; (void)lpCmd; (void)nShow;

    {
        WSADATA wsa;
        WSAStartup(MAKEWORD(2, 2), &wsa);
    }

    memset(&wc, 0, sizeof(wc));
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"RTL8152BNetCfg";

    if (!RegisterClass(&wc)) return -1;

    g_hWnd = CreateWindow(L"RTL8152BNetCfg", L"RTL8152B Net Config",
        WS_VISIBLE | WS_CAPTION | WS_SYSMENU | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX,
        0, 0, 480, 272, NULL, NULL, hInst, NULL);

    if (!g_hWnd) return -1;

    autoStarted = 0;

    if (lpCmd && lpCmd[0]) {
        if (wcsstr(lpCmd, L"/autostatic") || wcsstr(lpCmd, L"-autostatic")) {
            AddLog(L"AutoStart: autostatic requested");
            CreateThread(NULL, 0, AutoStartupThread,
                         (LPVOID)NETCFG_STARTUP_MODE_STATIC, 0, NULL);
            autoStarted = 1;
        } else if (wcsstr(lpCmd, L"/autodhcp") || wcsstr(lpCmd, L"-autodhcp")) {
            AddLog(L"AutoStart: autodhcp requested");
            CreateThread(NULL, 0, AutoStartupThread,
                         (LPVOID)NETCFG_STARTUP_MODE_DHCP, 0, NULL);
            autoStarted = 1;
        } else if (wcsstr(lpCmd, L"/autonet") || wcsstr(lpCmd, L"-autonet")) {
            AddLog(L"AutoStart: autonet requested");
            if (InterlockedExchange(&g_bBusy, 1) == 0) {
                CreateThread(NULL, 0, NetCheckThread, NULL, 0, NULL);
                autoStarted = 1;
            }
        }
    }

    if (!autoStarted) {
#if NETCFG_DEFAULT_STARTUP_MODE == NETCFG_STARTUP_MODE_STATIC
        AddLog(L"AutoStart: default static mode");
        CreateThread(NULL, 0, AutoStartupThread,
                     (LPVOID)NETCFG_STARTUP_MODE_STATIC, 0, NULL);
#elif NETCFG_DEFAULT_STARTUP_MODE == NETCFG_STARTUP_MODE_DHCP
        AddLog(L"AutoStart: default DHCP mode");
        CreateThread(NULL, 0, AutoStartupThread,
                     (LPVOID)NETCFG_STARTUP_MODE_DHCP, 0, NULL);
#endif
    }

    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    WSACleanup();
    return (int)msg.wParam;
}



