/* ============================================================================
 * rtl8152.c  -  Realtek RTL8152B USB-to-Fast-Ethernet
 *               NDIS 5.1 Miniport + USB Client Driver  (WinCE 7.0 / ARM)
 *
 * Author      : SweatierUnicorn
 * GitHub      : https://github.com/SweatierUnicorn
 * Distribution: Provided free of charge by the author.
 *
 *
 * Target chip : RTL8152B (VID=0x0BDA PID=0x8152 REV=0x2000)
 * Target HW   : Panasonic Strada CN-F1X10BD (Renesas ARM (Renesas SoC R8A77400 (R-Mobile A1)), WEC7 Automotive)
 * Compiler    : VS2008 cl.exe x86_arm, C89
 * Deployment  : DmyClassDrv replacement (LoadClientsDummy)
 * ============================================================================ */

#include "rtl8152.h"

/* ============================================================================
 * Logging - ring buffer for early boot, flush to \MediaCard\ when ready
 * ============================================================================ */
static WCHAR g_LogPath[64]  = {0};
static BOOL  g_LogPathInit  = FALSE;
static BOOL  g_MediaCardOk  = FALSE;

#define BOOT_LOG_SIZE 16384
static char  g_BootLog[BOOT_LOG_SIZE];
static int   g_BootLogPos    = 0;
static BOOL  g_BootLogFlushed = FALSE;

static void LogInitPath(void)
{
    int n;
    if (g_LogPathInit) return;
    for (n = 1; n <= 99; n++) {
        wsprintfW(g_LogPath, L"\\MediaCard\\rtl8152_%d.log", n);
        if (GetFileAttributes(g_LogPath) == 0xFFFFFFFF)
            break;
    }
    g_LogPathInit = TRUE;
}

static void LogMsg(const WCHAR *msg)
{
    HANDLE hFile;
    DWORD written;
    WCHAR buf[512];
    SYSTEMTIME st;
    char abuf[1024];
    int len;

    GetLocalTime(&st);
    wsprintfW(buf, L"[%02d:%02d:%02d.%03d] %s",
              st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, msg);
    OutputDebugString(buf);
    OutputDebugString(L"\r\n");

    len = 0;
    while (buf[len] && len < 1020) { abuf[len] = (char)buf[len]; len++; }
    abuf[len] = '\r';  abuf[len+1] = '\n';
    len += 2;

    if (!g_MediaCardOk) {
        if (GetFileAttributes(L"\\MediaCard\\") != 0xFFFFFFFF)
            g_MediaCardOk = TRUE;
    }
    if (g_MediaCardOk) {
        LogInitPath();
        hFile = CreateFile(g_LogPath, GENERIC_WRITE, FILE_SHARE_READ,
                           NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
        if (hFile != INVALID_HANDLE_VALUE) {
            SetFilePointer(hFile, 0, NULL, FILE_END);
            if (!g_BootLogFlushed && g_BootLogPos > 0) {
                WriteFile(hFile, g_BootLog, g_BootLogPos, &written, NULL);
                g_BootLogFlushed = TRUE;
            }
            WriteFile(hFile, abuf, len, &written, NULL);
            FlushFileBuffers(hFile);
            CloseHandle(hFile);
            return;
        }
        g_MediaCardOk = FALSE;
    }
    if (!g_BootLogFlushed && (g_BootLogPos + len < BOOT_LOG_SIZE)) {
        memcpy(g_BootLog + g_BootLogPos, abuf, len);
        g_BootLogPos += len;
    }
}

static void LogFlushBootBuffer(void)
{
    HANDLE hFile;
    DWORD written;
    if (g_BootLogFlushed || g_BootLogPos <= 0) return;
    if (!g_MediaCardOk) {
        if (GetFileAttributes(L"\\MediaCard\\") == 0xFFFFFFFF) return;
        g_MediaCardOk = TRUE;
    }
    LogInitPath();
    hFile = CreateFile(g_LogPath, GENERIC_WRITE, FILE_SHARE_READ,
                       NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) { g_MediaCardOk = FALSE; return; }
    SetFilePointer(hFile, 0, NULL, FILE_END);
    WriteFile(hFile, g_BootLog, g_BootLogPos, &written, NULL);
    FlushFileBuffers(hFile);
    CloseHandle(hFile);
    g_BootLogFlushed = TRUE;
}

/* ============================================================================
 * Global State
 * ============================================================================ */
static HINSTANCE    g_hDllInst        = NULL;
static NDIS_HANDLE  g_NdisWrapperHandle = NULL;
static BOOL         g_NdisRegistered  = FALSE;
static BOOL         g_RegKeysCreated  = FALSE;
static volatile LONG g_bAttached       = 0;

/* USB state (populated by USBDeviceAttach) */
static USB_HANDLE   g_UsbDevice       = NULL;
static LPCUSB_FUNCS g_UsbFuncs        = NULL;
static USB_PIPE     g_BulkInPipe      = NULL;
static USB_PIPE     g_BulkOutPipe     = NULL;
static USB_PIPE     g_InterruptPipe   = NULL;

/* Adapter instance (created by RtlInitialize) */
static PRTL8152_ADAPTER g_Adapter     = NULL;

/* Prevent concurrent NdisInitThread instances */
static volatile LONG g_NdisInitRunning = 0;

/* Link monitor thread control */
static volatile LONG g_LinkMonitorRunning = 0;

/* Packet filter & lookahead from OID sets */
static ULONG g_PacketFilter = 0;
static ULONG g_LookAhead    = 1500;

/* Statistics (stubs for Step 1) */
static volatile ULONG g_TxOk   = 0;
static volatile ULONG g_TxFail = 0;
static volatile ULONG g_RxPkts = 0;
static volatile ULONG g_RxErr  = 0;
static volatile ULONG g_RxThreadState = 0;
static volatile ULONG g_RxThreadLoops = 0;
static volatile DWORD g_RxThreadBeat  = 0;
static volatile DWORD g_RxLastBytesRead = 0;
static volatile DWORD g_RxLastOffset = 0;
static volatile DWORD g_RxLastOpts1 = 0;
static volatile DWORD g_RxLastPktLen = 0;
static volatile LONG  g_RxIndicateReady = 0;

/* ============================================================================
 * Dynamic NDIS Loading
 *
 * k.usbd.dll calls LoadDriver("rtl8152.dll"). If rtl8152.dll statically
 * imports NDIS.dll, the PE loader fails to resolve the dependency and
 * LoadDriver returns NULL. All Panasonic USB client drivers import ONLY COREDLL.
 *
 * Solution: load NDIS functions at runtime via GetProcAddress.
 * ============================================================================ */
static HMODULE g_hNdisDll = NULL;

/* --- NDIS function typedefs --- */
typedef VOID (NDISAPI *PFN_NdisInitializeWrapper)(
    OUT PNDIS_HANDLE, IN PVOID, IN PVOID, IN PVOID);
typedef NDIS_STATUS (NDISAPI *PFN_NdisMRegisterMiniport)(
    IN NDIS_HANDLE, IN PNDIS_MINIPORT_CHARACTERISTICS, IN UINT);
typedef VOID (NDISAPI *PFN_NdisTerminateWrapper)(
    IN NDIS_HANDLE, IN PVOID);
typedef NDIS_STATUS (NDISAPI *PFN_NdisAllocateMemoryWithTag)(
    OUT PVOID *, IN UINT, IN ULONG);
typedef VOID (NDISAPI *PFN_NdisFreeMemory)(
    IN PVOID, IN UINT, IN UINT);
typedef VOID (NDISAPI *PFN_NdisMSetAttributesEx)(
    IN NDIS_HANDLE, IN NDIS_HANDLE, IN UINT, IN ULONG, IN NDIS_INTERFACE_TYPE);
typedef VOID (NDISAPI *PFN_NdisMIndicateStatus)(
    IN NDIS_HANDLE, IN NDIS_STATUS, IN PVOID, IN UINT);
typedef VOID (NDISAPI *PFN_NdisMIndicateStatusComplete)(
    IN NDIS_HANDLE);

/* --- NDIS function pointers (resolved at runtime) --- */
static PFN_NdisInitializeWrapper       pfnNdisInitializeWrapper       = NULL;
static PFN_NdisMRegisterMiniport       pfnNdisMRegisterMiniport       = NULL;
static PFN_NdisTerminateWrapper        pfnNdisTerminateWrapper        = NULL;
static PFN_NdisAllocateMemoryWithTag   pfnNdisAllocateMemoryWithTag   = NULL;
static PFN_NdisFreeMemory              pfnNdisFreeMemory              = NULL;
static PFN_NdisMSetAttributesEx        pfnNdisMSetAttributesEx        = NULL;
static PFN_NdisMIndicateStatus         pfnNdisMIndicateStatus         = NULL;
static PFN_NdisMIndicateStatusComplete pfnNdisMIndicateStatusComplete = NULL;

static BOOL LoadNdisFunctions(void)
{
    if (g_hNdisDll) return TRUE;
    g_hNdisDll = LoadLibrary(L"ndis.dll");
    if (!g_hNdisDll) {
        WCHAR em[96];
        wsprintfW(em, L"LoadNdis: LoadLibrary(ndis.dll) FAIL err=%d", GetLastError());
        LogMsg(em);
        return FALSE;
    }
    LogMsg(L"LoadNdis: ndis.dll loaded OK");

    pfnNdisInitializeWrapper = (PFN_NdisInitializeWrapper)
        GetProcAddressW(g_hNdisDll, L"NdisInitializeWrapper");
    pfnNdisMRegisterMiniport = (PFN_NdisMRegisterMiniport)
        GetProcAddressW(g_hNdisDll, L"NdisMRegisterMiniport");
    pfnNdisTerminateWrapper = (PFN_NdisTerminateWrapper)
        GetProcAddressW(g_hNdisDll, L"NdisTerminateWrapper");
    pfnNdisAllocateMemoryWithTag = (PFN_NdisAllocateMemoryWithTag)
        GetProcAddressW(g_hNdisDll, L"NdisAllocateMemoryWithTag");
    pfnNdisFreeMemory = (PFN_NdisFreeMemory)
        GetProcAddressW(g_hNdisDll, L"NdisFreeMemory");
    pfnNdisMSetAttributesEx = (PFN_NdisMSetAttributesEx)
        GetProcAddressW(g_hNdisDll, L"NdisMSetAttributesEx");
    pfnNdisMIndicateStatus = (PFN_NdisMIndicateStatus)
        GetProcAddressW(g_hNdisDll, L"NdisMIndicateStatus");
    pfnNdisMIndicateStatusComplete = (PFN_NdisMIndicateStatusComplete)
        GetProcAddressW(g_hNdisDll, L"NdisMIndicateStatusComplete");

    if (!pfnNdisInitializeWrapper || !pfnNdisMRegisterMiniport ||
        !pfnNdisTerminateWrapper || !pfnNdisAllocateMemoryWithTag ||
        !pfnNdisFreeMemory || !pfnNdisMSetAttributesEx ||
        !pfnNdisMIndicateStatus || !pfnNdisMIndicateStatusComplete) {
        WCHAR em[128];
        wsprintfW(em, L"LoadNdis: MISSING funcs: IW=%d MR=%d TW=%d AM=%d FM=%d SA=%d IS=%d IC=%d",
            pfnNdisInitializeWrapper?1:0, pfnNdisMRegisterMiniport?1:0,
            pfnNdisTerminateWrapper?1:0, pfnNdisAllocateMemoryWithTag?1:0,
            pfnNdisFreeMemory?1:0, pfnNdisMSetAttributesEx?1:0,
            pfnNdisMIndicateStatus?1:0, pfnNdisMIndicateStatusComplete?1:0);
        LogMsg(em);
        FreeLibrary(g_hNdisDll);
        g_hNdisDll = NULL;
        return FALSE;
    }
    LogMsg(L"LoadNdis: all 8 functions resolved OK");
    return TRUE;
}

/* ============================================================================
 * Forward Declarations
 * ============================================================================ */
static NDIS_STATUS RtlInitialize(
    OUT PNDIS_STATUS OpenErrorStatus,
    OUT PUINT SelectedMediumIndex,
    IN PNDIS_MEDIUM MediumArray,
    IN UINT MediumArraySize,
    IN NDIS_HANDLE MiniportAdapterHandle,
    IN NDIS_HANDLE WrapperConfigurationContext);
static VOID RtlHalt(IN NDIS_HANDLE MiniportAdapterContext);
static NDIS_STATUS RtlReset(OUT PBOOLEAN AddressingReset,
                             IN NDIS_HANDLE MiniportAdapterContext);
static NDIS_STATUS RtlSend(IN NDIS_HANDLE MiniportAdapterContext,
                            IN PNDIS_PACKET Packet, IN UINT Flags);
static NDIS_STATUS RtlQueryInformation(
    IN NDIS_HANDLE MiniportAdapterContext, IN NDIS_OID Oid,
    IN PVOID InformationBuffer, IN ULONG InformationBufferLength,
    OUT PULONG BytesWritten, OUT PULONG BytesNeeded);
static NDIS_STATUS RtlSetInformation(
    IN NDIS_HANDLE MiniportAdapterContext, IN NDIS_OID Oid,
    IN PVOID InformationBuffer, IN ULONG InformationBufferLength,
    OUT PULONG BytesRead, OUT PULONG BytesNeeded);
static BOOLEAN RtlCheckForHang(IN NDIS_HANDLE MiniportAdapterContext);
static VOID RtlShutdown(IN PVOID ShutdownContext);

static void CreateNdisRegistryKeys(void);
static void AppendTcpipBind(void);
static BOOL WaitForNetCfgInstanceId(WCHAR *guidBuf, DWORD cchGuidBuf);
static void EnsureTcpipInterfaceKey(void);
static void RebindTcpipAdapter(HANDLE hNdis);
static DWORD WINAPI NdisInitThread(LPVOID param);
static DWORD WINAPI RxThread(LPVOID param);

typedef struct _RTL8152_TX_DESC {
    DWORD opts1;
    DWORD opts2;
} RTL8152_TX_DESC;

typedef struct _RTL8152_RX_DESC {
    DWORD opts1;
    DWORD opts2;
    DWORD opts3;
    DWORD opts4;
    DWORD opts5;
    DWORD opts6;
} RTL8152_RX_DESC;

static NDIS_STATUS RtlBulkOutSync(PVOID buffer, DWORD length);
static NDIS_STATUS RtlBulkInSync(PVOID buffer, DWORD bufferSize, DWORD *bytesRead);
static NDIS_STATUS RtlCopyPacketToBuffer(PNDIS_PACKET Packet, UCHAR *buffer,
                                         UINT bufferSize, UINT *packetLength);
static NDIS_STATUS WriteMacAddress(const UCHAR mac[6]);
static NDIS_STATUS ApplyRxPacketFilter(ULONG packetFilter);

/* ============================================================================
 * USB Control Transfers for RTL8152B (synchronous)
 *
 * RTL8152B USB vendor transfer protocol (Linux r8152.c):
 *
 * READ:  bmRequestType=0xC0, bRequest=0x05
 *        wValue = register address (DWORD-aligned for 1/2/4 byte reads)
 *        wIndex = MCU_TYPE only (no byte-enables for reads)
 *        wLength = 4 always (for 1/2/4 byte reads); 6 for MAC address
 *        Result: extract byte/word from the 4-byte response
 *
 * WRITE: bmRequestType=0x40, bRequest=0x05
 *        wValue = register address, DWORD-aligned (reg & ~3)
 *        wIndex = MCU_TYPE | byte-enable bits:
 *          BYTE write at byte N:  MCU_TYPE | (BYTE_EN_BYTE << N)  [N=reg&3]
 *          WORD write at lower:   MCU_TYPE | BYTE_EN_WORD         [reg&2==0]
 *          WORD write at upper:   MCU_TYPE | (BYTE_EN_DWORD << 8) [reg&2!=0]
 *          DWORD write:           MCU_TYPE | BYTE_EN_DWORD
 *        wLength = 4 always
 *        Data: 4-byte LE buffer, value shifted to correct byte position
 * ============================================================================ */

/* Helper: issue one USB vendor transfer, wait, get status, close.
 * Returns NDIS_STATUS_SUCCESS only if hXfer != NULL and dwErr == USB_NO_ERROR.
 * isRead: non-zero = IN transfer, zero = OUT transfer. */
static NDIS_STATUS
DoVendorTransfer(USHORT wValue, USHORT wIndex, USHORT wLength,
                 PVOID buf, int isRead)
{
    USB_DEVICE_REQUEST req;
    USB_TRANSFER hXfer;
    DWORD dwBytes = 0, dwErr = 0;
    DWORD timeout = 0;

    if (!g_UsbFuncs || !g_UsbDevice) return NDIS_STATUS_FAILURE;

    req.bmRequestType = isRead ? RTL8152_REQT_READ : RTL8152_REQT_WRITE;
    req.bRequest      = RTL8152_REQ_GET_REGS;   /* 0x05 for both R/W */
    req.wValue        = wValue;
    req.wIndex        = wIndex;
    req.wLength       = wLength;

    hXfer = g_UsbFuncs->lpIssueVendorTransfer(
        g_UsbDevice, NULL, NULL,
        isRead ? (USB_IN_TRANSFER | USB_SHORT_TRANSFER_OK) : USB_OUT_TRANSFER,
        &req, buf, 0);
    if (!hXfer) {
        WCHAR em[96];
        wsprintfW(em, L"USB_%s: hXfer=NULL wi=%04X reg=%04X sz=%d",
                  isRead ? L"RD" : L"WR", wIndex, wValue, (int)wLength);
        LogMsg(em);
        return NDIS_STATUS_FAILURE;
    }

    while (!g_UsbFuncs->lpIsTransferComplete(hXfer)) {
        Sleep(1);
        if (++timeout > 2000) {
            WCHAR em[64];
            wsprintfW(em, L"USB_%s: TIMEOUT reg=%04X", isRead ? L"RD" : L"WR", wValue);
            LogMsg(em);
            g_UsbFuncs->lpAbortTransfer(hXfer, 0);
            g_UsbFuncs->lpCloseTransfer(hXfer);
            return NDIS_STATUS_FAILURE;
        }
    }
    g_UsbFuncs->lpGetTransferStatus(hXfer, &dwBytes, &dwErr);
    g_UsbFuncs->lpCloseTransfer(hXfer);
    if (dwErr != USB_NO_ERROR) {
        WCHAR em[96];
        wsprintfW(em, L"USB_%s: ERR=%d wi=%04X reg=%04X",
                  isRead ? L"RD" : L"WR", dwErr, wIndex, wValue);
        LogMsg(em);
    }
    return (dwErr == USB_NO_ERROR) ? NDIS_STATUS_SUCCESS : NDIS_STATUS_FAILURE;
}

static NDIS_STATUS
RtlBulkOutSync(PVOID buffer, DWORD length)
{
    USB_TRANSFER hXfer;
    DWORD dwBytes = 0, dwErr = 0;
    DWORD timeout = 0;

    if (!g_UsbFuncs || !g_BulkOutPipe || !buffer || !length)
        return NDIS_STATUS_FAILURE;

    hXfer = g_UsbFuncs->lpIssueBulkTransfer(
        g_BulkOutPipe, NULL, NULL,
        USB_OUT_TRANSFER,
        length, buffer, 0);
    if (!hXfer) {
        WCHAR em[80];
        wsprintfW(em, L"TX: IssueBulkTransfer NULL len=%u", length);
        LogMsg(em);
        return NDIS_STATUS_FAILURE;
    }

    while (!g_UsbFuncs->lpIsTransferComplete(hXfer)) {
        Sleep(1);
        if (++timeout > 3000) {
            LogMsg(L"TX: bulk timeout");
            g_UsbFuncs->lpAbortTransfer(hXfer, 0);
            g_UsbFuncs->lpCloseTransfer(hXfer);
            return NDIS_STATUS_FAILURE;
        }
    }

    g_UsbFuncs->lpGetTransferStatus(hXfer, &dwBytes, &dwErr);
    g_UsbFuncs->lpCloseTransfer(hXfer);

    if (dwErr != USB_NO_ERROR || dwBytes != length) {
        WCHAR em[96];
        wsprintfW(em, L"TX: bulk err=%u bytes=%u/%u", dwErr, dwBytes, length);
        LogMsg(em);
        return NDIS_STATUS_FAILURE;
    }

    return NDIS_STATUS_SUCCESS;
}

static NDIS_STATUS
RtlBulkInSync(PVOID buffer, DWORD bufferSize, DWORD *bytesRead)
{
    USB_TRANSFER hXfer;
    DWORD dwBytes = 0, dwErr = 0;
    DWORD timeout = 0;
    WCHAR em[96];

    if (bytesRead) *bytesRead = 0;
    if (!g_UsbFuncs || !g_BulkInPipe || !buffer || !bufferSize)
        return NDIS_STATUS_FAILURE;

    hXfer = g_UsbFuncs->lpIssueBulkTransfer(
        g_BulkInPipe, NULL, NULL,
        USB_IN_TRANSFER | USB_SHORT_TRANSFER_OK,
        bufferSize, buffer, 0);
    if (!hXfer) {
        LogMsg(L"RX: lpIssueBulkTransfer returned NULL");
        return NDIS_STATUS_FAILURE;
    }

    while (!g_UsbFuncs->lpIsTransferComplete(hXfer)) {
        Sleep(1);
        if (++timeout > 3000) {
            g_UsbFuncs->lpAbortTransfer(hXfer, 0);
            g_UsbFuncs->lpCloseTransfer(hXfer);
            LogMsg(L"RX: bulk IN timeout");
            return NDIS_STATUS_FAILURE;
        }
    }

    g_UsbFuncs->lpGetTransferStatus(hXfer, &dwBytes, &dwErr);
    g_UsbFuncs->lpCloseTransfer(hXfer);

    if (bytesRead) *bytesRead = dwBytes;
    if (dwErr != USB_NO_ERROR) {
        wsprintfW(em, L"RX: bulk IN status err=%u bytes=%u", dwErr, dwBytes);
        LogMsg(em);
        return NDIS_STATUS_FAILURE;
    }

    return NDIS_STATUS_SUCCESS;
}

static NDIS_STATUS
RtlCopyPacketToBuffer(PNDIS_PACKET Packet, UCHAR *buffer,
                      UINT bufferSize, UINT *packetLength)
{
    PNDIS_BUFFER currentBuffer;
    UINT copied = 0;

    if (!Packet || !buffer || !packetLength)
        return NDIS_STATUS_FAILURE;

    currentBuffer = Packet->Private.Head;
    if (!currentBuffer)
        return NDIS_STATUS_FAILURE;

    while (currentBuffer) {
        PVOID virtualAddress = NULL;
        UINT length = 0;

        NdisQueryBuffer(currentBuffer, &virtualAddress, &length);
        if (!virtualAddress || copied + length > bufferSize)
            return NDIS_STATUS_FAILURE;

        memcpy(buffer + copied, virtualAddress, length);
        copied += length;
        NdisGetNextBuffer(currentBuffer, &currentBuffer);
    }

    if (copied == 0)
        return NDIS_STATUS_FAILURE;

    *packetLength = copied;
    return NDIS_STATUS_SUCCESS;
}

static DWORD WINAPI RxThread(LPVOID param)
{
    PRTL8152_ADAPTER A = (PRTL8152_ADAPTER)param;
    UCHAR *rxBuf = NULL;
    ULONG baseAddr;
    UCHAR *alignedBuf;
    DWORD emptyReads = 0;

    g_RxThreadState = 1;
    g_RxThreadBeat = GetTickCount();

    rxBuf = (UCHAR*)LocalAlloc(0, RTL_AGG_BUF_SIZE + RTL_RX_ALIGN);
    if (!rxBuf) {
        g_RxThreadState = 0xEE;
        g_RxThreadBeat = GetTickCount();
        LogMsg(L"RX: thread alloc FAILED");
        if (A) A->hRxThread = NULL;
        return 1;
    }

    baseAddr = (ULONG)rxBuf;
    alignedBuf = (UCHAR*)(((baseAddr + (RTL_RX_ALIGN - 1)) / RTL_RX_ALIGN) * RTL_RX_ALIGN);

    g_RxThreadState = 2;
    g_RxThreadBeat = GetTickCount();
    LogMsg(L"RX: thread started");
    g_RxThreadState = 3;
    g_RxThreadBeat = GetTickCount();

    while (A && !A->bExitRxThread && !A->bUsbGone &&
           InterlockedCompareExchange(&g_bAttached, 1, 1) == 1) {
        DWORD bytesRead = 0;
        DWORD offset = 0;
        NDIS_STATUS st;
        g_RxThreadState = 4;
        g_RxThreadLoops++;
        g_RxThreadBeat = GetTickCount();
        st = RtlBulkInSync(alignedBuf, RTL_AGG_BUF_SIZE, &bytesRead);

        if (st != NDIS_STATUS_SUCCESS) {
            g_RxThreadState = 5;
            g_RxThreadBeat = GetTickCount();
            Sleep(10);
            continue;
        }

        if (bytesRead == 0) {
            g_RxThreadState = 6;
            g_RxThreadBeat = GetTickCount();
            emptyReads++;
            if ((emptyReads & 0xFF) == 0) {
                LogMsg(L"RX: bulk IN completed with 0 bytes");
            }
            Sleep(1);
            continue;
        }

        g_RxThreadState = 7;
        g_RxThreadBeat = GetTickCount();
        g_RxLastBytesRead = bytesRead;
        g_RxLastOffset = 0;
        g_RxLastOpts1 = 0;
        g_RxLastPktLen = 0;
        emptyReads = 0;

        while (bytesRead >= offset + RTL_RX_DESC_SIZE) {
            RTL8152_RX_DESC rxDesc;
            DWORD opts1;
            DWORD pktLen;
            UCHAR *frame;
            UINT lookaheadLen;
            WCHAR rm[80];

            g_RxThreadState = 8;
            g_RxThreadBeat = GetTickCount();
            g_RxLastOffset = offset;
            memcpy(&rxDesc, alignedBuf + offset, sizeof(rxDesc));
            opts1 = rxDesc.opts1;
            pktLen = opts1 & RX_LEN_MASK;
            g_RxLastOpts1 = opts1;
            g_RxLastPktLen = pktLen;
            g_RxThreadState = 9;
            g_RxThreadBeat = GetTickCount();

            if (pktLen < 64 || pktLen > RTL_MAX_ETH_SIZE + ETH_FCS_LEN) {
                g_RxThreadState = 10;
                g_RxErr++;
                break;
            }

            if (bytesRead < offset + RTL_RX_DESC_SIZE + pktLen) {
                g_RxThreadState = 11;
                g_RxErr++;
                break;
            }

            if (opts1 & RD_CRC) {
                g_RxThreadState = 12;
                g_RxErr++;
            } else {
                frame = alignedBuf + offset + RTL_RX_DESC_SIZE;
                if (pktLen > ETH_FCS_LEN)
                    pktLen -= ETH_FCS_LEN;

                if (pktLen >= ETH_HEADER_SIZE && A->MiniportAdapterHandle) {
                    g_RxThreadState = 13;
                    if (InterlockedCompareExchange(&g_RxIndicateReady, 1, 1) == 1) {
                        lookaheadLen = (pktLen > ETH_HEADER_SIZE) ? (UINT)(pktLen - ETH_HEADER_SIZE) : 0;
                        NdisMEthIndicateReceive(A->MiniportAdapterHandle,
                                                NULL,
                                                frame,
                                                ETH_HEADER_SIZE,
                                                frame + ETH_HEADER_SIZE,
                                                lookaheadLen,
                                                pktLen - ETH_HEADER_SIZE);
                        NdisMEthIndicateReceiveComplete(A->MiniportAdapterHandle);
                        g_RxThreadState = 14;
                        g_RxPkts++;
                        wsprintfW(rm, L"RX: packet OK len=%u", pktLen);
                        LogMsg(rm);
                    } else {
                        g_RxThreadState = 16;
                    }
                }
            }

            offset += RTL_RX_DESC_SIZE + pktLen + ETH_FCS_LEN;
            offset = (offset + (RTL_RX_ALIGN - 1)) & ~(RTL_RX_ALIGN - 1);
            g_RxThreadState = 15;
            g_RxThreadBeat = GetTickCount();
        }
    }

    g_RxThreadState = 8;
    g_RxThreadBeat = GetTickCount();
    LogMsg(L"RX: thread exit");
    LocalFree(rxBuf);
    if (A) A->hRxThread = NULL;
    return 0;
}

/* RtlReadReg - always reads 4 bytes (DWORD-aligned); extracts byte/word/dword.
 * Special case: size>4 (e.g. 6 for MAC) → direct transfer, no alignment. */
static NDIS_STATUS
RtlReadReg(USHORT mcuType, USHORT regOffset, USHORT size, PVOID data)
{
    DWORD dwordBuf = 0;
    NDIS_STATUS st;
    int wordShift;

    if (size > 4) {
        /* Direct transfer (MAC = 6 bytes): send as-is */
        return DoVendorTransfer(regOffset, mcuType, size, data, 1);
    }

    /* For 1/2/4 byte reads: always fetch 4 bytes from DWORD-aligned address */
    st = DoVendorTransfer(
        (USHORT)(regOffset & (USHORT)(~3)),   /* DWORD-aligned addr */
        mcuType,                               /* no byte-enables for reads */
        4,                                     /* always 4 bytes */
        &dwordBuf, 1);
    if (st != NDIS_STATUS_SUCCESS) return st;

    switch (size) {
    case 1:
        *(UCHAR*)data = (UCHAR)((dwordBuf >> ((regOffset & 3) * 8)) & 0xFF);
        break;
    case 2:
        wordShift = (regOffset & 2) ? 16 : 0;
        *(USHORT*)data = (USHORT)((dwordBuf >> wordShift) & 0xFFFF);
        break;
    default: /* 4 */
        *(DWORD*)data = dwordBuf;
        break;
    }
    return NDIS_STATUS_SUCCESS;
}

/* RtlWriteReg - always writes 4 bytes with byte-enables in wIndex.
 * Mirrors Linux r8152.c: ocp_write_byte/word/dword. */
static NDIS_STATUS
RtlWriteReg(USHORT mcuType, USHORT regOffset, USHORT size, PVOID data)
{
    DWORD dwordBuf = 0;
    USHORT byteen;
    USHORT alignedOffset = (USHORT)(regOffset & (USHORT)(~3));
    int shift;

    switch (size) {
    case 1:
        shift = regOffset & 3;
        dwordBuf = (DWORD)(*(UCHAR*)data) << (shift * 8);
        byteen   = (USHORT)(mcuType | ((USHORT)BYTE_EN_BYTE << shift));
        break;
    case 2:
        shift = regOffset & 2;
        dwordBuf = (DWORD)(*(USHORT*)data);
        byteen   = (USHORT)(mcuType | ((USHORT)BYTE_EN_WORD << shift));
        if (shift)
            dwordBuf <<= (shift * 8);
        break;
    default: /* 4 */
        dwordBuf = *(DWORD*)data;
        byteen   = (USHORT)(mcuType | BYTE_EN_DWORD);
        break;
    }

    return DoVendorTransfer(alignedOffset, byteen, 4, &dwordBuf, 0);
}

/* Convenience: read/write single byte/word/dword from PLA space */
static NDIS_STATUS RtlReadByte(USHORT reg, UCHAR *val)
{
    return RtlReadReg(MCU_TYPE_PLA, reg, 1, val);
}

static NDIS_STATUS RtlWriteByte(USHORT reg, UCHAR val)
{
    return RtlWriteReg(MCU_TYPE_PLA, reg, 1, &val);
}

static NDIS_STATUS RtlReadWord(USHORT reg, USHORT *val)
{
    return RtlReadReg(MCU_TYPE_PLA, reg, 2, val);
}

static NDIS_STATUS RtlWriteWord(USHORT reg, USHORT val)
{
    return RtlWriteReg(MCU_TYPE_PLA, reg, 2, &val);
}

static NDIS_STATUS RtlReadDword(USHORT reg, DWORD *val)
{
    return RtlReadReg(MCU_TYPE_PLA, reg, 4, val);
}

static NDIS_STATUS RtlWriteDword(USHORT reg, DWORD val)
{
    return RtlWriteReg(MCU_TYPE_PLA, reg, 4, &val);
}

/* ============================================================================
 * USB-space register helpers (MCU_TYPE_USB)
 * ============================================================================ */
static NDIS_STATUS RtlReadWordUsb(USHORT reg, USHORT *val)
{
    return RtlReadReg(MCU_TYPE_USB, reg, 2, val);
}

static NDIS_STATUS RtlWriteByteUsb(USHORT reg, UCHAR val)
{
    return RtlWriteReg(MCU_TYPE_USB, reg, 1, &val);
}

static NDIS_STATUS RtlWriteWordUsb(USHORT reg, USHORT val)
{
    return RtlWriteReg(MCU_TYPE_USB, reg, 2, &val);
}

static NDIS_STATUS RtlWriteDwordUsb(USHORT reg, DWORD val)
{
    return RtlWriteReg(MCU_TYPE_USB, reg, 4, &val);
}

/* ============================================================================
 * Read-Modify-Write helpers (PLA space)
 * ============================================================================ */
static NDIS_STATUS RtlSetWordBits(USHORT reg, USHORT bits)
{
    USHORT v = 0;
    if (RtlReadWord(reg, &v) != NDIS_STATUS_SUCCESS) return NDIS_STATUS_FAILURE;
    return RtlWriteWord(reg, v | bits);
}

static NDIS_STATUS RtlClrWordBits(USHORT reg, USHORT bits)
{
    USHORT v = 0;
    if (RtlReadWord(reg, &v) != NDIS_STATUS_SUCCESS) return NDIS_STATUS_FAILURE;
    return RtlWriteWord(reg, v & (USHORT)(~bits));
}

static NDIS_STATUS RtlSetByteBits(USHORT reg, UCHAR bits)
{
    UCHAR v = 0;
    if (RtlReadByte(reg, &v) != NDIS_STATUS_SUCCESS) return NDIS_STATUS_FAILURE;
    return RtlWriteByte(reg, v | bits);
}

/* USB-space Read-Modify-Write */
static NDIS_STATUS RtlSetWordBitsUsb(USHORT reg, USHORT bits)
{
    USHORT v = 0;
    if (RtlReadWordUsb(reg, &v) != NDIS_STATUS_SUCCESS) return NDIS_STATUS_FAILURE;
    return RtlWriteWordUsb(reg, v | bits);
}

static NDIS_STATUS RtlClrWordBitsUsb(USHORT reg, USHORT bits)
{
    USHORT v = 0;
    if (RtlReadWordUsb(reg, &v) != NDIS_STATUS_SUCCESS) return NDIS_STATUS_FAILURE;
    return RtlWriteWordUsb(reg, v & (USHORT)(~bits));
}

/* ============================================================================
 * PHY (MDIO) access via OCP GPHY mechanism
 *
 * Linux r8152_phy_write(addr, data):
 *   ocp_base = addr & 0xf000  → write to PLA_OCP_GPHY_BASE
 *   ocp_index = (addr & 0x0fff) | 0xb000  → write data to this PLA reg
 *
 * r8152_mdio_write(mii_reg, val):
 *   addr = OCP_BASE_MII + mii_reg * 2  (e.g. BMCR reg0 → 0xa400)
 * ============================================================================ */
static USHORT g_ocpBase = 0xFFFF;  /* cached base, 0xFFFF = not initialized */

static NDIS_STATUS RtlPhyReadWord(USHORT addr, USHORT *val)
{
    USHORT ocp_base  = addr & 0xF000;
    USHORT ocp_index = (addr & 0x0FFF) | 0xB000;
    if (ocp_base != g_ocpBase) {
        if (RtlWriteWord(PLA_OCP_GPHY_BASE, ocp_base) != NDIS_STATUS_SUCCESS)
            return NDIS_STATUS_FAILURE;
        g_ocpBase = ocp_base;
    }
    return RtlReadWord(ocp_index, val);
}

static NDIS_STATUS RtlPhyWriteWord(USHORT addr, USHORT val)
{
    USHORT ocp_base  = addr & 0xF000;
    USHORT ocp_index = (addr & 0x0FFF) | 0xB000;
    if (ocp_base != g_ocpBase) {
        if (RtlWriteWord(PLA_OCP_GPHY_BASE, ocp_base) != NDIS_STATUS_SUCCESS)
            return NDIS_STATUS_FAILURE;
        g_ocpBase = ocp_base;
    }
    return RtlWriteWord(ocp_index, val);
}

/* mdio_write: mii_reg 0-31 → addr = OCP_BASE_MII + reg*2 */
static NDIS_STATUS RtlMdioWrite(USHORT mii_reg, USHORT val)
{
    return RtlPhyWriteWord((USHORT)(OCP_BASE_MII + mii_reg * 2), val);
}

static NDIS_STATUS RtlMdioRead(USHORT mii_reg, USHORT *val)
{
    return RtlPhyReadWord((USHORT)(OCP_BASE_MII + mii_reg * 2), val);
}

/* mdio_clr_bit: read-modify-write PHY register */
static NDIS_STATUS RtlMdioClrBit(USHORT mii_reg, USHORT bit)
{
    USHORT v = 0;
    if (RtlMdioRead(mii_reg, &v) != NDIS_STATUS_SUCCESS) return NDIS_STATUS_FAILURE;
    return RtlMdioWrite(mii_reg, v & (USHORT)(~bit));
}

static NDIS_STATUS RtlMdioSetBit(USHORT mii_reg, USHORT bit)
{
    USHORT v = 0;
    if (RtlMdioRead(mii_reg, &v) != NDIS_STATUS_SUCCESS) return NDIS_STATUS_FAILURE;
    return RtlMdioWrite(mii_reg, v | bit);
}

/* OCP PHY write (for ALDPS config, addr like 0x2010):
 * same mechanism as phy_write but different address space */
static NDIS_STATUS RtlOcpPhyWrite(USHORT addr, USHORT val)
{
    return RtlPhyWriteWord(addr, val);
}

/* ============================================================================
 * RtlInitChip - chip initialization sequence for RTL8152B (VER_01)
 *
 * Mirrors Linux r8152.c:
 *   r8152b_init()       → power/clock setup
 *   r8152b_hw_phy_cfg() → PHY ALDPS + flow control
 *   rtl_enable()        → PLA_CR RE+TE, ungating RX DMA
 *
 * Called from USBDeviceAttach after pipes are valid.
 * ============================================================================ */
static void RtlInitChip(void)
{
    USHORT val16 = 0;
    DWORD  val32 = 0;

    LogMsg(L"INIT: r8152b_init start");

    /* 1. PHY power-up: clear BMCR.PDOWN (MII reg 0)
     *    Linux: r8152_mdio_test_and_clr_bit(tp, MII_BMCR, BMCR_PDOWN) */
    RtlMdioClrBit(MII_BMCR, BMCR_PDOWN);
    LogMsg(L"INIT: BMCR PDOWN cleared");

    /* 2. Disable ALDPS (Auto-Link Down Power Save)
     *    Linux: r8152_aldps_en(tp, false)
     *    ocp_reg_write(OCP_ALDPS_CONFIG, ENPDNPS | LINKENA | DIS_SDSAVE) + sleep 20ms */
    RtlOcpPhyWrite(OCP_ALDPS_CONFIG, ENPDNPS | LINKENA | DIS_SDSAVE);
    Sleep(20);
    LogMsg(L"INIT: ALDPS disabled");

    /* 3. RTL_VER_01 specific: clear LED_MODE bits in PLA_LED_FEATURE
     *    Linux: ocp_word_clr_bits(PLA_LED_FEATURE, LED_MODE_MASK) */
    RtlClrWordBits(PLA_LED_FEATURE, LED_MODE_MASK);
    LogMsg(L"INIT: LED_FEATURE cleared");

    /* 4. Disable power cut: clear USB_UPS_CTRL.POWER_CUT
     *    Linux: r8152_power_cut_en(tp, false)
     *    ocp_word_clr_bits(USB, USB_UPS_CTRL, POWER_CUT)
     *    ocp_word_clr_bits(USB, USB_PM_CTRL_STATUS, RESUME_INDICATE) */
    RtlClrWordBitsUsb(USB_UPS_CTRL, POWER_CUT);
    RtlClrWordBitsUsb(USB_PM_CTRL_STATUS, RESUME_INDICATE);
    LogMsg(L"INIT: power cut disabled");

    /* 5. PHY power register: set TX_10M_IDLE_EN and PFM_PWM_SWITCH
     *    Linux: ocp_word_set_bits(PLA, PLA_PHY_PWR, TX_10M_IDLE_EN | PFM_PWM_SWITCH) */
    RtlSetWordBits(PLA_PHY_PWR, TX_10M_IDLE_EN | PFM_PWM_SWITCH);
    LogMsg(L"INIT: PLA_PHY_PWR set");

    /* 6. MCU clock gating: (val & ~MCU_CLK_RATIO_MASK) | (MCU_CLK_RATIO | D3_CLK_GATED_EN)
     *    Linux: ocp_dword_w0w1(PLA, PLA_MAC_PWR_CTRL, MCU_CLK_RATIO_MASK,
     *                          MCU_CLK_RATIO | D3_CLK_GATED_EN) */
    if (RtlReadDword(PLA_MAC_PWR_CTRL, &val32) == NDIS_STATUS_SUCCESS) {
        val32 = (val32 & ~MCU_CLK_RATIO_MASK) | (MCU_CLK_RATIO | D3_CLK_GATED_EN);
        RtlWriteDword(PLA_MAC_PWR_CTRL, val32);
    }
    LogMsg(L"INIT: MAC_PWR_CTRL set");

    /* 7. GPHY interrupt mask
     *    Linux: ocp_write_word(PLA, PLA_GPHY_INTR_IMR,
     *           GPHY_STS_MSK | SPEED_DOWN_MSK | SPDWN_RXDV_MSK | SPDWN_LINKCHG_MSK) */
    RtlWriteWord(PLA_GPHY_INTR_IMR,
                 GPHY_STS_MSK | SPEED_DOWN_MSK | SPDWN_RXDV_MSK | SPDWN_LINKCHG_MSK);
    LogMsg(L"INIT: GPHY_INTR_IMR set");

    /* 8-10. USB timer sequence
     *    Linux: set bit15, write MSC_TIMER, clear bit15 */
    RtlSetWordBitsUsb(USB_USB_TIMER, 0x8000);
    RtlWriteWordUsb(USB_MSC_TIMER, 1000);  /* 8000ms / 8 = 1000 */
    RtlClrWordBitsUsb(USB_USB_TIMER, 0x8000);
    LogMsg(L"INIT: USB timer set");

    /* 11. Enable RX aggregation: clear RX_AGG_DISABLE | RX_ZERO_EN
     *    Linux: ocp_word_clr_bits(USB, USB_USB_CTRL, RX_AGG_DISABLE | RX_ZERO_EN) */
    RtlClrWordBitsUsb(USB_USB_CTRL, RX_AGG_DISABLE | RX_ZERO_EN);
    LogMsg(L"INIT: RX agg enabled");

    /* 11a. Basic TX/RX thresholds from Linux rtl8152_init_common */
    RtlWriteByteUsb(USB_TX_AGG, TX_AGG_MAX_THRESHOLD);
    RtlWriteDwordUsb(USB_RX_BUF_TH, RX_THR_HIGH);
    RtlWriteDwordUsb(USB_TX_DMA, TEST_MODE_DISABLE | TX_SIZE_ADJUST1);
    RtlWriteWord(PLA_RMS, RTL8152_RMS);
    RtlSetWordBits(PLA_TCR0, TCR0_AUTO_FIFO);
    LogMsg(L"INIT: TX/RX thresholds configured");

    /* ---- r8152b_hw_phy_cfg ---- */
    LogMsg(L"INIT: hw_phy_cfg start");

    /* 12. Re-enable ALDPS (with power save)
     *    Linux: r8152_aldps_en(tp, true)
     *    ocp_reg_write(OCP_ALDPS_CONFIG, ENPWRSAVE | ENPDNPS | LINKENA | DIS_SDSAVE) */
    RtlOcpPhyWrite(OCP_ALDPS_CONFIG, ENPWRSAVE | ENPDNPS | LINKENA | DIS_SDSAVE);
    LogMsg(L"INIT: ALDPS re-enabled");

    /* 13. Flow control advertisement: MII_ADVERTISE |= PAUSE_CAP | PAUSE_ASYM
     *    Linux: r8152b_enable_fc → r8152_mdio_set_bit(MII_ADVERTISE,
     *           ADVERTISE_PAUSE_CAP | ADVERTISE_PAUSE_ASYM) */
    RtlMdioSetBit(MII_ANAR, ADVERTISE_PAUSE_CAP | ADVERTISE_PAUSE_ASYM);
    LogMsg(L"INIT: flow control set");

    /* ---- rtl_enable (start RX/TX) ---- */
    LogMsg(L"INIT: rtl_enable start");

    /* 14. Reset packet filter: PLA_FMC clr then set FMC_FCR_MCU_EN
     *    Linux: r8152b_reset_packet_filter */
    RtlClrWordBits(PLA_FMC, FMC_FCR_MCU_EN);
    RtlSetWordBits(PLA_FMC, FMC_FCR_MCU_EN);
    LogMsg(L"INIT: packet filter reset");

    /* 15. Enable RX and TX: PLA_CR |= CR_RE | CR_TE
     *    Linux: ocp_byte_set_bits(PLA, PLA_CR, CR_RE | CR_TE) */
    RtlSetByteBits(PLA_CR, CR_RE | CR_TE);
    LogMsg(L"INIT: PLA_CR RE+TE enabled");

    /* 16. Ungate RX DMA: PLA_MISC_1 &= ~RXDY_GATED_EN
     *    Linux: rxdy_gated_en(tp, false) */
    RtlClrWordBits(PLA_MISC_1, RXDY_GATED_EN);
    LogMsg(L"INIT: RXDY gate removed");

    /* Read back PHYSTATUS to see if link came up */
    {
        UCHAR phyStat = 0;
        if (RtlReadByte(PLA_PHYSTATUS, &phyStat) == NDIS_STATUS_SUCCESS) {
            WCHAR pm[64];
            wsprintfW(pm, L"INIT: PHYSTATUS=%02X link=%s",
                      phyStat, (phyStat & LINK_STATUS) ? L"UP" : L"DOWN (need cable?)");
            LogMsg(pm);
        }
    }

    /* Read back PLA_CR to confirm */
    {
        UCHAR cr = 0;
        if (RtlReadByte(PLA_CR, &cr) == NDIS_STATUS_SUCCESS) {
            WCHAR cm[48];
            wsprintfW(cm, L"INIT: PLA_CR=0x%02X", cr);
            LogMsg(cm);
        }
    }

    LogMsg(L"INIT: RtlInitChip DONE");
}

/* ============================================================================
 * Chip Version Detection
 * Read PLA_TCR0, extract (val >> 16) & 0x7CF0 to identify chip revision.
 * ============================================================================ */
static UCHAR DetectChipVersion(void)
{
    DWORD tcr0 = 0;
    USHORT ver;
    WCHAR em[80];

    if (RtlReadDword(PLA_TCR0, &tcr0) != NDIS_STATUS_SUCCESS) {
        LogMsg(L"RTL: TCR0 read FAIL, assuming VER_01");
        return RTL_VER_01;
    }

    ver = (USHORT)((tcr0 >> 16) & VERSION_MASK);
    wsprintfW(em, L"RTL: TCR0=0x%08X ver_field=0x%04X", tcr0, ver);
    LogMsg(em);

    switch (ver) {
    case 0x4C00: return RTL_VER_01;
    case 0x4C10: return RTL_VER_02;
    case 0x5C00: return RTL_VER_03;
    case 0x5C10: return RTL_VER_04;
    case 0x5C20: return RTL_VER_05;
    default:
        wsprintfW(em, L"RTL: unknown version 0x%04X, assuming VER_01", ver);
        LogMsg(em);
        return RTL_VER_01;
    }
}

/* ============================================================================
 * Read MAC Address from PLA_IDR (0xC000)
 * Must enable config write mode (CRWECR=0xC0) before reading,
 * then restore to normal (CRWECR=0x00) after.
 * ============================================================================ */
static NDIS_STATUS ReadMacAddress(UCHAR mac[6])
{
    NDIS_STATUS st;

    /* Enable config mode */
    RtlWriteByte(PLA_CRWECR, CRWECR_CONFIG);
    Sleep(1);

    /* Read 6-byte MAC from PLA_IDR */
    st = RtlReadReg(MCU_TYPE_PLA, PLA_IDR, 6, mac);

    /* Restore normal mode */
    RtlWriteByte(PLA_CRWECR, CRWECR_NORAML);

    return st;
}

static NDIS_STATUS ApplyRxPacketFilter(ULONG packetFilter)
{
    DWORD rcr = RCR_APM | RCR_AB;
    DWORD mar[2];
    NDIS_STATUS st;
    WCHAR pm[96];

    mar[0] = 0;
    mar[1] = 0;

    if (packetFilter & NDIS_PACKET_TYPE_PROMISCUOUS) {
        rcr |= RCR_AAP | RCR_AM;
        mar[0] = 0xFFFFFFFFUL;
        mar[1] = 0xFFFFFFFFUL;
    } else {
        if (packetFilter & (NDIS_PACKET_TYPE_MULTICAST | NDIS_PACKET_TYPE_ALL_MULTICAST)) {
            rcr |= RCR_AM;
            mar[0] = 0xFFFFFFFFUL;
            mar[1] = 0xFFFFFFFFUL;
        }
    }

    st = DoVendorTransfer(PLA_MAR,
                          (USHORT)(MCU_TYPE_PLA | BYTE_EN_DWORD),
                          sizeof(mar),
                          mar,
                          0);
    if (st != NDIS_STATUS_SUCCESS) {
        LogMsg(L"OID: ApplyRxPacketFilter MAR write FAILED");
        return st;
    }

    st = RtlWriteDword(PLA_RCR, rcr);
    if (st != NDIS_STATUS_SUCCESS) {
        LogMsg(L"OID: ApplyRxPacketFilter RCR write FAILED");
        return st;
    }

    wsprintfW(pm, L"OID: hw RX filter RCR=%08X NDIS=%08X", rcr, packetFilter);
    LogMsg(pm);
    return NDIS_STATUS_SUCCESS;
}

static NDIS_STATUS WriteMacAddress(const UCHAR mac[6])
{
    UCHAR macBuf[8];
    NDIS_STATUS st;

    if (!mac)
        return NDIS_STATUS_FAILURE;

    memset(macBuf, 0, sizeof(macBuf));
    memcpy(macBuf, mac, 6);

    RtlWriteByte(PLA_CRWECR, CRWECR_CONFIG);
    Sleep(1);
    st = DoVendorTransfer(PLA_IDR,
                          (USHORT)(MCU_TYPE_PLA | BYTE_EN_SIX_BYTES),
                          8,
                          macBuf,
                          0);
    RtlWriteByte(PLA_CRWECR, CRWECR_NORAML);
    return st;
}

/* ============================================================================
 * DllMain
 * ============================================================================ */
BOOL WINAPI DllMain(HINSTANCE hInst, DWORD dwReason, LPVOID lpRes)
{
    (void)lpRes;
    if (dwReason == DLL_PROCESS_ATTACH) {
        g_hDllInst = hInst;
        DisableThreadLibraryCalls(hInst);
        LogMsg(L"DllMain: ATTACH (RTL8152B)");
    }
    return TRUE;
}

/* ============================================================================
 * LinkMonitorThread - polls PHYSTATUS every 1 second.
 * Calls NdisMIndicateStatus(MEDIA_CONNECT/DISCONNECT) on state changes.
 * Exits automatically when the USB device detaches (g_bAttached becomes 0).
 * ============================================================================ */
static DWORD WINAPI LinkMonitorThread(LPVOID param)
{
    UCHAR phyStat  = 0;
    BOOL  wasUp    = FALSE;
    BOOL  haveState = FALSE;
    DWORD interval = 0;
    DWORD lastRxDiagTick = 0;
    DWORD rxExitCode = 0;

    (void)param;
    LogMsg(L"LinkMon: thread started");

    while (InterlockedCompareExchange(&g_bAttached, 1, 1) == 1) {
        /* Read PHYSTATUS (PLA register 0xE908) */
        phyStat = 0;
        if (RtlReadByte(PLA_PHYSTATUS, &phyStat) == NDIS_STATUS_SUCCESS) {
            BOOL isUp = (phyStat & LINK_STATUS) ? TRUE : FALSE;

            if (!haveState || isUp != wasUp) {
                WCHAR pm[80];
                wsprintfW(pm, L"LinkMon: link %s (PHYSTATUS=%02X)",
                          isUp ? L"UP" : L"DOWN", (DWORD)phyStat);
                LogMsg(pm);

                if (g_Adapter &&
                    pfnNdisMIndicateStatus &&
                    pfnNdisMIndicateStatusComplete &&
                    g_Adapter->MiniportAdapterHandle) {
                    if (isUp) {
                        g_Adapter->CurrentLinkSpeed = 100000000UL / 100; /* 100Mbps */
                        pfnNdisMIndicateStatus(g_Adapter->MiniportAdapterHandle,
                                               NDIS_STATUS_MEDIA_CONNECT, NULL, 0);
                    } else {
                        pfnNdisMIndicateStatus(g_Adapter->MiniportAdapterHandle,
                                               NDIS_STATUS_MEDIA_DISCONNECT, NULL, 0);
                    }
                    pfnNdisMIndicateStatusComplete(g_Adapter->MiniportAdapterHandle);
                    g_Adapter->bMediaConnected = isUp;
                    wasUp = isUp;
                    haveState = TRUE;
                }
            }
        }

        if (g_Adapter && g_Adapter->hRxThread) {
            DWORD now = GetTickCount();
            if (now - lastRxDiagTick >= 2000) {
                WCHAR rm[192];
                rxExitCode = 0xFFFFFFFFUL;
                GetExitCodeThread(g_Adapter->hRxThread, &rxExitCode);
                wsprintfW(rm, L"LinkMon: RX state=%u loops=%u beat=%u bytes=%u off=%u opts1=%08X pkt=%u pkts=%u err=%u exit=%u",
                          g_RxThreadState, g_RxThreadLoops, g_RxThreadBeat,
                          g_RxLastBytesRead, g_RxLastOffset,
                          g_RxLastOpts1, g_RxLastPktLen,
                          g_RxPkts, g_RxErr, rxExitCode);
                LogMsg(rm);
                lastRxDiagTick = now;
            }
        }

        /* Poll every 1 second when UP, 500ms when DOWN (faster detection) */
        interval = wasUp ? 1000 : 500;
        Sleep(interval);
    }

    LogMsg(L"LinkMon: thread exit (USB detached)");
    InterlockedExchange(&g_LinkMonitorRunning, 0);
    return 0;
}

/* ============================================================================
 * USBDeviceAttach - DmyClassDrv replacement
 *   Accept ALL devices (fAcceptControl=TRUE).
 *   Only do real work for Realtek RTL8152B (VID=0BDA, PID=8152).
 * ============================================================================ */
BOOL USBDeviceAttach(
    USB_HANDLE hDevice,
    LPCUSB_FUNCS lpUsbFuncs,
    LPCUSB_INTERFACE lpInterface,
    LPCWSTR szUniqueDriverId,
    LPBOOL fAcceptControl,
    LPCUSB_DRIVER_SETTINGS lpDriverSettings,
    DWORD dwUnused)
{
    DWORD i;
    LPCUSB_DEVICE lpDevInfo;
    LPCUSB_INTERFACE pIf = NULL;
    UCHAR mac[6];
    UCHAR chipVer;

    (void)szUniqueDriverId;
    (void)lpDriverSettings;
    (void)dwUnused;

    LogFlushBootBuffer();
    LogMsg(L"USB: USBDeviceAttach ENTER");

    /* DmyClassDrv catch-all: always accept */
    *fAcceptControl = TRUE;

    /* Get device descriptor */
    lpDevInfo = lpUsbFuncs->lpGetDeviceInfo(hDevice);
    if (!lpDevInfo) {
        LogMsg(L"USB: lpGetDeviceInfo=NULL, accepting as DmyClassDrv");
        return TRUE;
    }

    /* Log every device */
    {
        WCHAR dm[128];
        wsprintfW(dm, L"USB: VID=%04X PID=%04X bcd=%04X",
                  lpDevInfo->Descriptor.idVendor,
                  lpDevInfo->Descriptor.idProduct,
                  lpDevInfo->Descriptor.bcdDevice);
        LogMsg(dm);
    }

    /* Only work with Realtek RTL8152B */
    if (lpDevInfo->Descriptor.idVendor != 0x0BDA ||
        lpDevInfo->Descriptor.idProduct != 0x8152) {
        return TRUE;  /* not our device — accept silently like DmyClassDrv */
    }

    LogMsg(L"=== USB: RTL8152B detected! ===");

    /* Guard: skip if already attached (prevents repeated init on re-enum) */
    if (InterlockedCompareExchange(&g_bAttached, 1, 0) != 0) {
        LogMsg(L"USB: already attached, skip");
        return TRUE;
    }

    /* ---- Find interface and endpoints ---- */
    g_UsbDevice = hDevice;
    g_UsbFuncs  = lpUsbFuncs;
    g_BulkInPipe  = NULL;
    g_BulkOutPipe = NULL;
    g_InterruptPipe = NULL;

    if (lpInterface != NULL) {
        pIf = (LPCUSB_INTERFACE)lpInterface;
        LogMsg(L"USB: lpInterface provided");
    } else {
        LPCUSB_CONFIGURATION pCfg;

        LogMsg(L"USB: lpInterface=NULL, searching manually");

        /* Try FindInterface with class=0xFF, subclass=0xFF (vendor-specific) */
        pIf = lpUsbFuncs->lpFindInterface(lpDevInfo, 0xFF, 0xFF);
        if (pIf) {
            LogMsg(L"USB: lpFindInterface(FF,FF) OK");
        } else {
            /* Fallback: active config, first interface */
            pCfg = lpDevInfo->lpActiveConfig;
            if (pCfg && pCfg->dwNumInterfaces > 0 && pCfg->lpInterfaces) {
                pIf = &pCfg->lpInterfaces[0];
                LogMsg(L"USB: using lpActiveConfig->lpInterfaces[0]");
            } else if (lpDevInfo->lpConfigs) {
                pCfg = &lpDevInfo->lpConfigs[0];
                if (pCfg && pCfg->dwNumInterfaces > 0 && pCfg->lpInterfaces) {
                    pIf = &pCfg->lpInterfaces[0];
                    LogMsg(L"USB: using lpConfigs[0]->lpInterfaces[0]");
                }
            }
        }
    }

    if (!pIf || !pIf->lpEndpoints) {
        LogMsg(L"USB: FAIL - no interface/endpoints found");
        g_UsbDevice = NULL;
        g_UsbFuncs = NULL;
        InterlockedExchange(&g_bAttached, 0);
        return TRUE;
    }

    {
        WCHAR em[80];
        wsprintfW(em, L"USB: interface has %d endpoints",
                  pIf->Descriptor.bNumEndpoints);
        LogMsg(em);
    }

    /* Open pipes */
    for (i = 0; i < pIf->Descriptor.bNumEndpoints; i++) {
        LPCUSB_ENDPOINT ep = &pIf->lpEndpoints[i];
        UCHAR attr = ep->Descriptor.bmAttributes & USB_ENDPOINT_TYPE_MASK;

        {
            WCHAR em[96];
            wsprintfW(em, L"USB: EP[%d] addr=%02X attr=%02X maxPkt=%d",
                      i, ep->Descriptor.bEndpointAddress,
                      ep->Descriptor.bmAttributes,
                      ep->Descriptor.wMaxPacketSize);
            LogMsg(em);
        }

        if (attr == USB_ENDPOINT_TYPE_BULK) {
            if (USB_ENDPOINT_DIRECTION_IN(ep->Descriptor.bEndpointAddress)) {
                g_BulkInPipe = lpUsbFuncs->lpOpenPipe(hDevice, &ep->Descriptor);
                {
                    WCHAR em[64];
                    wsprintfW(em, L"USB: Bulk IN pipe=%08X", (DWORD)g_BulkInPipe);
                    LogMsg(em);
                }
            } else {
                g_BulkOutPipe = lpUsbFuncs->lpOpenPipe(hDevice, &ep->Descriptor);
                {
                    WCHAR em[64];
                    wsprintfW(em, L"USB: Bulk OUT pipe=%08X", (DWORD)g_BulkOutPipe);
                    LogMsg(em);
                }
            }
        } else if (attr == USB_ENDPOINT_TYPE_INTERRUPT) {
            g_InterruptPipe = lpUsbFuncs->lpOpenPipe(hDevice, &ep->Descriptor);
            {
                WCHAR em[64];
                wsprintfW(em, L"USB: Interrupt pipe=%08X", (DWORD)g_InterruptPipe);
                LogMsg(em);
            }
        }
    }

    if (!g_BulkInPipe || !g_BulkOutPipe) {
        LogMsg(L"USB: FAIL - missing bulk pipes");
        g_UsbDevice = NULL;
        g_UsbFuncs = NULL;
        InterlockedExchange(&g_bAttached, 0);
        return TRUE;
    }

    /* ---- RTL8152B chip init ---- */

    /* Step 1: Detect chip version */
    chipVer = DetectChipVersion();
    {
        WCHAR vm[64];
        wsprintfW(vm, L"USB: ChipVersion=%d (RTL_VER_%02d)", chipVer, chipVer);
        LogMsg(vm);
    }

    /* Step 2: Read MAC address from PLA_IDR */
    memset(mac, 0, sizeof(mac));
    if (ReadMacAddress(mac) == NDIS_STATUS_SUCCESS) {
        WCHAR mm[80];
        wsprintfW(mm, L"USB: MAC=%02X:%02X:%02X:%02X:%02X:%02X",
                  mac[0], mac[1], mac[2], mac[3], mac[4], mac[5]);
        LogMsg(mm);
    } else {
        LogMsg(L"USB: MAC read FAIL");
    }

    /* Step 3: Chip initialization (r8152b_init + phy_cfg + rtl_enable) */
    RtlInitChip();

    /* Do not indicate RX into NDIS until the stack reapplies packet filter
     * after attach. Early indication during attach/rebind currently crashes
     * inside NdisMEthIndicateReceive on this Panasonic WinCE build. */
    InterlockedExchange(&g_RxIndicateReady, 0);

    /* Step 3a: Program a safe default RX filter before traffic starts.
     * WinCE typically requests 0x0000000B shortly after bind, but until then
     * hardware must already accept own unicast and broadcast frames. */
    ApplyRxPacketFilter(g_PacketFilter ? g_PacketFilter : (NDIS_PACKET_TYPE_DIRECTED | NDIS_PACKET_TYPE_BROADCAST));

    /* Step 4: Keep hardware MAC aligned with the MAC already cached by NDIS.
     * WinCE boot scan calls RtlInitialize before USB attach, so upper layers may
     * already be using the fallback MAC. If hardware stays on the factory MAC,
     * replies will be addressed to a different destination MAC and RX will miss them. */
    if (g_Adapter) {
        if (WriteMacAddress(g_Adapter->MacAddress) == NDIS_STATUS_SUCCESS) {
            WCHAR mm[96];
            wsprintfW(mm, L"USB: hardware MAC synced to NDIS MAC %02X:%02X:%02X:%02X:%02X:%02X",
                      g_Adapter->MacAddress[0], g_Adapter->MacAddress[1],
                      g_Adapter->MacAddress[2], g_Adapter->MacAddress[3],
                      g_Adapter->MacAddress[4], g_Adapter->MacAddress[5]);
            LogMsg(mm);
            g_Adapter->bHasMac = TRUE;
        } else {
            LogMsg(L"USB: hardware MAC sync FAILED");
        }
    }

    /* Step 5: Read PHY status after init */
    {
        UCHAR phyStat = 0;
        if (RtlReadByte(PLA_PHYSTATUS, &phyStat) == NDIS_STATUS_SUCCESS) {
            WCHAR pm[64];
            wsprintfW(pm, L"USB: PHYSTATUS=%02X link=%s",
                      phyStat, (phyStat & LINK_STATUS) ? L"UP" : L"DOWN");
            LogMsg(pm);
        }
    }

    /* Step 6: Start link monitor thread (polls PHYSTATUS, indicates MEDIA_CONNECT) */
    if (InterlockedExchange(&g_LinkMonitorRunning, 1) == 0) {
        LogMsg(L"USB: starting LinkMonitorThread");
        CreateThread(NULL, 0, LinkMonitorThread, NULL, 0, NULL);
    } else {
        LogMsg(L"USB: LinkMonitorThread already running");
    }

    if (g_Adapter && !g_Adapter->hRxThread) {
        WCHAR tm[96];
        DWORD exitCode = 0;
        g_Adapter->bExitRxThread = FALSE;
        g_Adapter->bUsbGone = FALSE;
        LogMsg(L"USB: starting RxThread");
        g_Adapter->hRxThread = CreateThread(NULL, 0x20000, RxThread, g_Adapter, 0, NULL);
        wsprintfW(tm, L"USB: RxThread handle=%08X err=%u",
                  (DWORD)g_Adapter->hRxThread, GetLastError());
        LogMsg(tm);
        if (g_Adapter->hRxThread) {
            Sleep(20);
            if (GetExitCodeThread(g_Adapter->hRxThread, &exitCode)) {
                wsprintfW(tm, L"USB: RxThread exitCode=%u", exitCode);
                LogMsg(tm);
            } else {
                wsprintfW(tm, L"USB: RxThread GetExitCodeThread FAIL err=%u", GetLastError());
                LogMsg(tm);
            }
        }
    }

    /* ---- Start NDIS registration thread ---- */
    if (InterlockedExchange(&g_NdisInitRunning, 1) == 0) {
        LogMsg(L"USB: starting NdisInitThread");
        CreateThread(NULL, 0, NdisInitThread, NULL, 0, NULL);
    } else {
        LogMsg(L"USB: NdisInitThread already running");
    }

    LogMsg(L"USB: USBDeviceAttach EXIT");
    return TRUE;
}

/* ============================================================================
 * NdisInitThread - background thread to register adapter with NDIS
 * ============================================================================ */
static DWORD WINAPI NdisInitThread(LPVOID param)
{
    int waitSec;
    HANDLE hNdis;
    BOOL bResult;
    DWORD cbOut;
    WCHAR msz[24];
    DWORD cbIn;

    (void)param;

    Sleep(500);
    LogFlushBootBuffer();
    LogMsg(L"NdisInit: start");

    /* Create registry keys for NDIS */
    CreateNdisRegistryKeys();

    /* Wait for NDS0: (NDIS device) */
    hNdis = INVALID_HANDLE_VALUE;
    for (waitSec = 0; waitSec < 25; waitSec++) {
        {
            WCHAR ww[96];
            wsprintfW(ww, L"NdisInit: tick %d/25 NdisReg=%d Adapter=%08X",
                      waitSec, g_NdisRegistered, (DWORD)g_Adapter);
            LogMsg(ww);
        }

        if (g_NdisRegistered) {
            LogMsg(L"NdisInit: registered by boot scan!");
            goto post_register;
        }

        hNdis = CreateFile(L"NDS0:", GENERIC_READ | GENERIC_WRITE,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           NULL, OPEN_ALWAYS, 0, NULL);
        if (hNdis != INVALID_HANDLE_VALUE) {
            WCHAR ww[64];
            wsprintfW(ww, L"NdisInit: NDS0: found at sec %d", waitSec);
            LogMsg(ww);
            CloseHandle(hNdis);
            hNdis = INVALID_HANDLE_VALUE;
            break;
        }
        Sleep(1000);
    }
    if (waitSec >= 25) {
        LogMsg(L"NdisInit: NDS0: NOT found in 25s, abort");
        goto done;
    }

    /* Let NDIS boot scan finish */
    Sleep(2000);

    if (g_NdisRegistered) {
        LogMsg(L"NdisInit: registered during stabilization");
        goto post_register;
    }

    LogMsg(L"NdisInit: NDIS did not call DriverEntry. Using IOCTL.");

    /* IOCTL_NDIS_REGISTER_ADAPTER (0x170026) */
    hNdis = CreateFile(L"NDS0:", GENERIC_READ | GENERIC_WRITE,
                       FILE_SHARE_READ | FILE_SHARE_WRITE,
                       NULL, OPEN_ALWAYS, 0, NULL);
    if (hNdis == INVALID_HANDLE_VALUE) {
        WCHAR em[64];
        wsprintfW(em, L"NdisInit: NDS0: reopen FAIL err=%d", GetLastError());
        LogMsg(em);
        goto done;
    }

    /* Multi_sz: "RTL8152B\0RTL8152B1\0\0" */
    memset(msz, 0, sizeof(msz));
    wcscpy(msz, L"RTL8152B");         /* [0..7]='RTL8152B', [8]='\0' */
    wcscpy(msz + 9, L"RTL8152B1");    /* [9..17]='RTL8152B1', [18]='\0' */
    /* msz[19] = '\0' from memset = multi_sz final null */
    cbIn = 20 * sizeof(WCHAR);
    cbOut = 0;

    LogMsg(L"NdisInit: IOCTL REGISTER_ADAPTER...");
    bResult = DeviceIoControl(hNdis, 0x170026,
                              msz, cbIn, NULL, 0, &cbOut, NULL);
    {
        WCHAR bm[128];
        wsprintfW(bm, L"NdisInit: REGISTER result=%d err=%d NdisReg=%d",
                  bResult, GetLastError(), g_NdisRegistered);
        LogMsg(bm);
    }

    /* IOCTL_NDIS_BIND_PROTOCOLS (0x170032) */
    Sleep(500);
    memset(msz, 0, sizeof(msz));
    wcscpy(msz, L"RTL8152B1");        /* [0..8]='RTL8152B1', [9]='\0' */
    wcscpy(msz + 10, L"TCPIP");       /* [10..14]='TCPIP', [15]='\0' */
    /* msz[16] = '\0' = multi_sz final null */
    cbIn = 17 * sizeof(WCHAR);
    cbOut = 0;

    LogMsg(L"NdisInit: IOCTL BIND_PROTOCOLS TCPIP...");
    bResult = DeviceIoControl(hNdis, 0x170032,
                              msz, cbIn, NULL, 0, &cbOut, NULL);
    {
        WCHAR bm[80];
        wsprintfW(bm, L"NdisInit: BIND result=%d err=%d", bResult, GetLastError());
        LogMsg(bm);
    }

post_register:
    if (hNdis == INVALID_HANDLE_VALUE) {
        hNdis = CreateFile(L"NDS0:", GENERIC_READ | GENERIC_WRITE,
                           FILE_SHARE_READ | FILE_SHARE_WRITE,
                           NULL, OPEN_ALWAYS, 0, NULL);
        if (hNdis == INVALID_HANDLE_VALUE) {
            WCHAR em[80];
            wsprintfW(em, L"NdisInit: post-setup NDS0 open FAIL err=%d",
                      GetLastError());
            LogMsg(em);
            goto done;
        }
    }

    EnsureTcpipInterfaceKey();
    AppendTcpipBind();
    RebindTcpipAdapter(hNdis);

    CloseHandle(hNdis);
    hNdis = INVALID_HANDLE_VALUE;

done:
    /* Try to flush boot log to MediaCard */
    if (!g_BootLogFlushed) {
        int w;
        for (w = 0; w < 15; w++) {
            Sleep(2000);
            if (GetFileAttributes(L"\\MediaCard\\") != 0xFFFFFFFF) {
                LogMsg(L"NdisInit: boot log flushed");
                break;
            }
        }
    }

    InterlockedExchange(&g_NdisInitRunning, 0);
    LogMsg(L"NdisInit: done");
    return 0;
}

/* ============================================================================
 * CreateNdisRegistryKeys - idempotent
 * Creates: Comm\RTL8152B, Comm\RTL8152B\Linkage, Comm\RTL8152B1,
 *          Comm\RTL8152B1\Parms, Comm\RTL8152B1\Parms\TcpIp
 * ============================================================================ */
static void CreateNdisRegistryKeys(void)
{
    HKEY hKey;
    DWORD dwDisp, dwVal;
    LONG rc;

    if (g_RegKeysCreated) return;
    LogMsg(L"CreateReg: start");

    /* Comm\RTL8152B (driver key) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B", 0, NULL, 0,
                        0, NULL, &hKey, &dwDisp);
    if (rc == ERROR_SUCCESS) {
        RegSetValueEx(hKey, L"DisplayName", 0, REG_SZ,
            (const BYTE*)L"Realtek RTL8152B USB Ethernet",
            30 * sizeof(WCHAR));
        RegSetValueEx(hKey, L"Group", 0, REG_SZ,
            (const BYTE*)L"NDIS", 5 * sizeof(WCHAR));
        RegSetValueEx(hKey, L"ImagePath", 0, REG_SZ,
            (const BYTE*)L"rtl8152.dll", sizeof(L"rtl8152.dll"));
        RegCloseKey(hKey);
        LogMsg(L"CreateReg: Comm\\RTL8152B OK");
    }

    /* Comm\RTL8152B\Linkage - Route = RTL8152B1 */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B\\Linkage", 0, NULL, 0,
                        0, NULL, &hKey, &dwDisp);
    if (rc == ERROR_SUCCESS) {
        WCHAR msz[16];
        wcscpy(msz, L"RTL8152B1");
        msz[9] = L'\0'; msz[10] = L'\0';
        RegSetValueEx(hKey, L"Route", 0, REG_MULTI_SZ,
            (const BYTE*)msz, 11 * sizeof(WCHAR));
        RegCloseKey(hKey);
        LogMsg(L"CreateReg: Linkage\\Route=RTL8152B1 OK");
    }

    /* Comm\RTL8152B1 (adapter instance) */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1", 0, NULL, 0,
                        0, NULL, &hKey, &dwDisp);
    if (rc == ERROR_SUCCESS) {
        RegSetValueEx(hKey, L"DisplayName", 0, REG_SZ,
            (const BYTE*)L"Realtek RTL8152B USB Ethernet",
            30 * sizeof(WCHAR));
        RegSetValueEx(hKey, L"Group", 0, REG_SZ,
            (const BYTE*)L"NDIS", 5 * sizeof(WCHAR));
        RegCloseKey(hKey);
        LogMsg(L"CreateReg: Comm\\RTL8152B1 OK");
    }

    /* Comm\RTL8152B1\Parms */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms", 0, NULL, 0,
                        0, NULL, &hKey, &dwDisp);
    if (rc == ERROR_SUCCESS) {
        dwVal = 0;
        RegSetValueEx(hKey, L"BusNumber", 0, REG_DWORD, (const BYTE*)&dwVal, 4);
        RegSetValueEx(hKey, L"BusType", 0, REG_DWORD, (const BYTE*)&dwVal, 4);
        dwVal = 6;  /* IF_TYPE_ETHERNET_CSMACD */
        RegSetValueEx(hKey, L"*IfType", 0, REG_DWORD, (const BYTE*)&dwVal, 4);
        dwVal = 0;  /* NdisMedium802_3 */
        RegSetValueEx(hKey, L"*MediaType", 0, REG_DWORD, (const BYTE*)&dwVal, 4);
        dwVal = 0;
        RegSetValueEx(hKey, L"*PhysicalMediaType", 0, REG_DWORD, (const BYTE*)&dwVal, 4);

        /* UpperBind = MULTI_SZ "TCPIP\0\0" */
        {
            WCHAR ub[8];
            memset(ub, 0, sizeof(ub));
            wcscpy(ub, L"TCPIP");
            RegSetValueEx(hKey, L"UpperBind", 0, REG_MULTI_SZ,
                (const BYTE*)ub, 7 * sizeof(WCHAR));
        }
        RegCloseKey(hKey);
        LogMsg(L"CreateReg: Parms OK");
    }

    /* Comm\RTL8152B1\Parms\TcpIp - static IP */
    rc = RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms\\TcpIp", 0,
                        NULL, 0, 0, NULL, &hKey, &dwDisp);
    if (rc == ERROR_SUCCESS) {
        dwVal = 0;
        RegSetValueEx(hKey, L"EnableDHCP", 0, REG_DWORD, (const BYTE*)&dwVal, 4);

        {
            WCHAR ipmsz[20];
            memset(ipmsz, 0, sizeof(ipmsz));
            wcscpy(ipmsz, L"192.168.50.100");
            RegSetValueEx(hKey, L"IpAddress", 0, REG_MULTI_SZ,
                (const BYTE*)ipmsz, (DWORD)((wcslen(ipmsz) + 2) * sizeof(WCHAR)));

            memset(ipmsz, 0, sizeof(ipmsz));
            wcscpy(ipmsz, L"255.255.255.0");
            RegSetValueEx(hKey, L"SubnetMask", 0, REG_MULTI_SZ,
                (const BYTE*)ipmsz, (DWORD)((wcslen(ipmsz) + 2) * sizeof(WCHAR)));

            memset(ipmsz, 0, sizeof(ipmsz));
            wcscpy(ipmsz, L"192.168.50.1");
            RegSetValueEx(hKey, L"DefaultGateway", 0, REG_MULTI_SZ,
                (const BYTE*)ipmsz, (DWORD)((wcslen(ipmsz) + 2) * sizeof(WCHAR)));
        }

        dwVal = 0;
        RegSetValueEx(hKey, L"AutoCfg", 0, REG_DWORD, (const BYTE*)&dwVal, 4);
        RegSetValueEx(hKey, L"IPAutoconfigurationEnabled", 0, REG_DWORD, (const BYTE*)&dwVal, 4);
        RegCloseKey(hKey);
        LogMsg(L"CreateReg: TcpIp OK");
    }

    g_RegKeysCreated = TRUE;
    LogMsg(L"CreateReg: done");
}

/* ============================================================================
 * AppendTcpipBind - add RTL8152B1 to Comm\Tcpip\Linkage\Bind
 * ============================================================================ */
static void AppendTcpipBind(void)
{
    HKEY hKey;
    DWORD dwDisp;

    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\Tcpip\\Linkage", 0,
                       NULL, 0, 0, NULL, &hKey, &dwDisp) == ERROR_SUCCESS) {
        BYTE existing[2048];
        DWORD cbData = sizeof(existing);
        DWORD dwType;

        if (RegQueryValueEx(hKey, L"Bind", NULL, &dwType,
                            existing, &cbData) == ERROR_SUCCESS
            && dwType == REG_MULTI_SZ && cbData > 4) {
            WCHAR *p = (WCHAR*)existing;
            int found = 0;
            while (*p) {
                if (wcscmp(p, L"RTL8152B1") == 0) { found = 1; break; }
                p += wcslen(p) + 1;
            }
            if (!found) {
                wcscpy(p, L"RTL8152B1");
                p += 10;
                *p = L'\0';
                cbData = (DWORD)((BYTE*)(p + 1) - existing);
                RegSetValueEx(hKey, L"Bind", 0, REG_MULTI_SZ, existing, cbData);
                LogMsg(L"TcpipBind: added RTL8152B1");
            } else {
                LogMsg(L"TcpipBind: already present");
            }
        } else {
            WCHAR msz[16];
            wcscpy(msz, L"RTL8152B1");
            msz[9] = L'\0'; msz[10] = L'\0';
            RegSetValueEx(hKey, L"Bind", 0, REG_MULTI_SZ,
                          (const BYTE*)msz, 11 * sizeof(WCHAR));
            LogMsg(L"TcpipBind: created with RTL8152B1");
        }
        RegCloseKey(hKey);
    } else {
        LogMsg(L"TcpipBind: cannot open Tcpip\\Linkage");
    }
}

static BOOL WaitForNetCfgInstanceId(WCHAR *guidBuf, DWORD cchGuidBuf)
{
    HKEY hKey;
    DWORD cb;
    DWORD dwType;
    LONG rc;
    int attempt;

    if (!guidBuf || cchGuidBuf == 0)
        return FALSE;

    guidBuf[0] = L'\0';

    for (attempt = 0; attempt < 20; attempt++) {
        rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, L"Comm\\RTL8152B1\\Parms",
                          0, KEY_READ, &hKey);
        if (rc == ERROR_SUCCESS) {
            cb = cchGuidBuf * sizeof(WCHAR);
            dwType = 0;
            rc = RegQueryValueEx(hKey, L"NetCfgInstanceId", NULL, &dwType,
                                 (BYTE*)guidBuf, &cb);
            RegCloseKey(hKey);
            if (rc == ERROR_SUCCESS && dwType == REG_SZ && guidBuf[0] == L'{') {
                WCHAR gm[96];
                guidBuf[cchGuidBuf - 1] = L'\0';
                wsprintfW(gm, L"NdisInit: NetCfgInstanceId=%s", guidBuf);
                LogMsg(gm);
                return TRUE;
            }
        }

        if ((attempt % 5) == 4)
            LogMsg(L"NdisInit: waiting for NetCfgInstanceId...");
        Sleep(1000);
    }

    LogMsg(L"NdisInit: NetCfgInstanceId not available");
    return FALSE;
}

static void EnsureTcpipInterfaceKey(void)
{
    WCHAR guidBuf[64];
    WCHAR ifPath[128];
    HKEY hKey;
    DWORD dwDisp;

    if (!WaitForNetCfgInstanceId(guidBuf, sizeof(guidBuf) / sizeof(guidBuf[0])))
        return;

    wsprintfW(ifPath, L"Comm\\Tcpip\\Parms\\Interfaces\\%s", guidBuf);
    if (RegCreateKeyEx(HKEY_LOCAL_MACHINE, ifPath, 0, NULL, 0,
                       0, NULL, &hKey, &dwDisp) == ERROR_SUCCESS) {
        RegCloseKey(hKey);
        LogMsg(L"NdisInit: Tcpip Interfaces marker OK");
    } else {
        LogMsg(L"NdisInit: Tcpip Interfaces marker FAIL");
    }
}

static void RebindTcpipAdapter(HANDLE hNdis)
{
    WCHAR msz[16];
    DWORD cbIn;
    DWORD cbOut;
    BOOL bResult;
    WCHAR bm[80];

    if (hNdis == INVALID_HANDLE_VALUE) {
        LogMsg(L"NdisInit: REBIND skip - NDS0 invalid");
        return;
    }

    memset(msz, 0, sizeof(msz));
    wcscpy(msz, L"RTL8152B1");
    cbIn = 11 * sizeof(WCHAR);
    cbOut = 0;
    bResult = DeviceIoControl(hNdis, 0x17002E,
                              msz, cbIn, NULL, 0, &cbOut, NULL);
    wsprintfW(bm, L"NdisInit: REBIND result=%d err=%d", bResult, GetLastError());
    LogMsg(bm);
}

/* ============================================================================
 * DriverEntry - called by NDIS (boot scan or IOCTL 0x170026)
 * ============================================================================ */
NTSTATUS DriverEntry(IN PVOID DriverObject, IN PVOID RegistryPath)
{
    NDIS_STATUS Status;
    NDIS_MINIPORT_CHARACTERISTICS RtlChar;

    LogFlushBootBuffer();

    {
        WCHAR dm[96];
        wsprintfW(dm, L"DriverEntry: DrvObj=%08X RegPath=%08X",
                  (DWORD)DriverObject, (DWORD)RegistryPath);
        LogMsg(dm);
    }

    if (g_NdisRegistered) {
        LogMsg(L"DriverEntry: already registered, skip");
        return NDIS_STATUS_SUCCESS;
    }

    if (!LoadNdisFunctions()) {
        LogMsg(L"DriverEntry: NDIS dynamic load FAILED");
        return NDIS_STATUS_FAILURE;
    }

    pfnNdisInitializeWrapper(
        &g_NdisWrapperHandle,
        DriverObject,
        RegistryPath,
        NULL);
    if (!g_NdisWrapperHandle) {
        LogMsg(L"DriverEntry: NdisInitializeWrapper FAILED");
        return NDIS_STATUS_FAILURE;
    }
    {
        WCHAR wm[64];
        wsprintfW(wm, L"DriverEntry: wrapper=%08X", (DWORD)g_NdisWrapperHandle);
        LogMsg(wm);
    }

    memset(&RtlChar, 0, sizeof(RtlChar));
    RtlChar.MajorNdisVersion        = 5;
    RtlChar.MinorNdisVersion        = 1;
    RtlChar.CheckForHangHandler     = RtlCheckForHang;
    RtlChar.InitializeHandler       = RtlInitialize;
    RtlChar.HaltHandler             = RtlHalt;
    RtlChar.ResetHandler            = RtlReset;
    RtlChar.SendHandler             = RtlSend;
    RtlChar.QueryInformationHandler = RtlQueryInformation;
    RtlChar.SetInformationHandler   = RtlSetInformation;

    /* NDIS 5.1 REQUIRES non-NULL AdapterShutdownHandler */
    RtlChar.AdapterShutdownHandler  = RtlShutdown;

    Status = pfnNdisMRegisterMiniport(
        g_NdisWrapperHandle, &RtlChar, sizeof(RtlChar));
    if (Status != NDIS_STATUS_SUCCESS) {
        WCHAR sm[64];
        wsprintfW(sm, L"DriverEntry: NdisMRegisterMiniport FAIL %08X", Status);
        LogMsg(sm);
        pfnNdisTerminateWrapper(g_NdisWrapperHandle, NULL);
        g_NdisWrapperHandle = NULL;
        return Status;
    }

    g_NdisRegistered = TRUE;
    LogMsg(L"DriverEntry: NDIS registered OK");
    return NDIS_STATUS_SUCCESS;
}

/* USB client stubs */
BOOL USBInstallDriver(LPCWSTR szDriverLibFile)
{
    (void)szDriverLibFile;
    return TRUE;
}

BOOL USBUnInstallDriver(void)
{
    return TRUE;
}

/* ============================================================================
 * RtlInitialize - called by NDIS after NdisMRegisterMiniport
 * ============================================================================ */
static NDIS_STATUS RtlInitialize(
    OUT PNDIS_STATUS OpenErrorStatus,
    OUT PUINT SelectedMediumIndex,
    IN PNDIS_MEDIUM MediumArray,
    IN UINT MediumArraySize,
    IN NDIS_HANDLE MiniportAdapterHandle,
    IN NDIS_HANDLE WrapperConfigurationContext)
{
    UINT i;
    PRTL8152_ADAPTER A;

    (void)OpenErrorStatus;
    (void)WrapperConfigurationContext;

    LogFlushBootBuffer();
    LogMsg(L"RtlInit: ENTERED");

    /* Find 802.3 medium */
    for (i = 0; i < MediumArraySize; i++) {
        if (MediumArray[i] == NdisMedium802_3) {
            *SelectedMediumIndex = i;
            break;
        }
    }
    if (i == MediumArraySize) {
        LogMsg(L"RtlInit: 802.3 medium not found!");
        return NDIS_STATUS_UNSUPPORTED_MEDIA;
    }

    /* Allocate adapter context */
    if (!pfnNdisAllocateMemoryWithTag ||
        pfnNdisAllocateMemoryWithTag((PVOID*)&A, sizeof(RTL8152_ADAPTER), 'LRTK')
        != NDIS_STATUS_SUCCESS) {
        LogMsg(L"RtlInit: alloc FAILED");
        return NDIS_STATUS_RESOURCES;
    }
    memset(A, 0, sizeof(RTL8152_ADAPTER));

    A->MiniportAdapterHandle = MiniportAdapterHandle;
    A->bMediaConnected = FALSE;
    A->bHasMac = FALSE;
    A->CurrentLinkSpeed = 1000000;  /* 100 Mbps in NDIS 100bps units */

    /* Fallback MAC (locally administered) */
    A->MacAddress[0] = 0x02;
    A->MacAddress[1] = 0x0B;
    A->MacAddress[2] = 0xDA;
    A->MacAddress[3] = 0x81;
    A->MacAddress[4] = 0x52;
    A->MacAddress[5] = 0x01;

    /* Register adapter attributes with NDIS */
    pfnNdisMSetAttributesEx(
        MiniportAdapterHandle,
        (NDIS_HANDLE)A,
        2,
        NDIS_ATTRIBUTE_DESERIALIZE | NDIS_ATTRIBUTE_USES_SAFE_BUFFER_APIS,
        0);

    g_Adapter = A;

    /* If USB already attached, try to read real MAC NOW.
     * NDIS queries OID_802_3_CURRENT_ADDRESS right after RtlInitialize
     * and caches the result. Must have real MAC before returning. */
    if (g_UsbDevice && g_UsbFuncs && g_BulkInPipe && g_BulkOutPipe) {
        if (ReadMacAddress(A->MacAddress) == NDIS_STATUS_SUCCESS) {
            A->bHasMac = TRUE;
            {
                WCHAR mm[80];
                wsprintfW(mm, L"RtlInit: MAC=%02X:%02X:%02X:%02X:%02X:%02X",
                          A->MacAddress[0], A->MacAddress[1], A->MacAddress[2],
                          A->MacAddress[3], A->MacAddress[4], A->MacAddress[5]);
                LogMsg(mm);
            }
        } else {
            LogMsg(L"RtlInit: early MAC read FAILED, using fallback");
        }
    } else {
        LogMsg(L"RtlInit: USB not ready, using fallback MAC");
    }

    /* Step 1: keep media disconnected, no HW init */
    pfnNdisMIndicateStatus(MiniportAdapterHandle,
                           NDIS_STATUS_MEDIA_DISCONNECT, NULL, 0);
    pfnNdisMIndicateStatusComplete(MiniportAdapterHandle);

    LogMsg(L"RtlInit: SUCCESS (Step1 - no TX/RX)");
    return NDIS_STATUS_SUCCESS;
}

/* ============================================================================
 * RtlHalt
 * ============================================================================ */
static VOID RtlHalt(IN NDIS_HANDLE MiniportAdapterContext)
{
    PRTL8152_ADAPTER A = (PRTL8152_ADAPTER)MiniportAdapterContext;
    LogMsg(L"RtlHalt: called");
    if (!A) return;
    A->bExitRxThread = TRUE;
    A->bUsbGone = TRUE;
    if (A->hRxThread) {
        WaitForSingleObject(A->hRxThread, 1000);
        CloseHandle(A->hRxThread);
        A->hRxThread = NULL;
    }
    if (pfnNdisFreeMemory)
        pfnNdisFreeMemory(A, sizeof(RTL8152_ADAPTER), 0);
    g_Adapter = NULL;
    g_NdisRegistered = FALSE;
}

/* ============================================================================
 * RtlReset / RtlCheckForHang / RtlShutdown
 * ============================================================================ */
static NDIS_STATUS RtlReset(OUT PBOOLEAN AddressingReset,
                             IN NDIS_HANDLE MiniportAdapterContext)
{
    (void)MiniportAdapterContext;
    *AddressingReset = FALSE;
    return NDIS_STATUS_SUCCESS;
}

static BOOLEAN RtlCheckForHang(IN NDIS_HANDLE MiniportAdapterContext)
{
    (void)MiniportAdapterContext;
    return FALSE;
}

static VOID RtlShutdown(IN PVOID ShutdownContext)
{
    (void)ShutdownContext;
    LogMsg(L"RtlShutdown: called");
}

/* ============================================================================
 * RtlSend - minimal synchronous TX path.
 * Builds one RTL8152B TX descriptor + Ethernet frame and sends it via Bulk OUT.
 * No checksum offload, no TSO, no TX aggregation yet.
 * ============================================================================ */
static NDIS_STATUS RtlSend(IN NDIS_HANDLE MiniportAdapterContext,
                            IN PNDIS_PACKET Packet, IN UINT Flags)
{
    PRTL8152_ADAPTER A = (PRTL8152_ADAPTER)MiniportAdapterContext;
    UCHAR txBuf[RTL_TX_DESC_SIZE + RTL_MAX_ETH_SIZE];
    RTL8152_TX_DESC *txDesc = (RTL8152_TX_DESC*)txBuf;
    UINT packetLength = 0;
    UINT sendLength;
    NDIS_STATUS st;
    WCHAR tm[80];

    (void)A;
    (void)Flags;

    if (!g_BulkOutPipe || !g_UsbFuncs || !g_UsbDevice) {
        LogMsg(L"TX: USB not ready");
        g_TxFail++;
        return NDIS_STATUS_FAILURE;
    }

    st = RtlCopyPacketToBuffer(Packet,
                               txBuf + RTL_TX_DESC_SIZE,
                               RTL_MAX_ETH_SIZE,
                               &packetLength);
    if (st != NDIS_STATUS_SUCCESS || packetLength == 0 || packetLength > RTL_MAX_ETH_SIZE) {
        LogMsg(L"TX: packet copy failed");
        g_TxFail++;
        return NDIS_STATUS_FAILURE;
    }

    if (packetLength < RTL_MIN_ETH_PAYLOAD) {
        memset(txBuf + RTL_TX_DESC_SIZE + packetLength, 0,
               RTL_MIN_ETH_PAYLOAD - packetLength);
        packetLength = RTL_MIN_ETH_PAYLOAD;
    }

    txDesc->opts1 = TX_FS | TX_LS | (packetLength & TX_LEN_MAX);
    txDesc->opts2 = 0;
    sendLength = RTL_TX_DESC_SIZE + packetLength;

    st = RtlBulkOutSync(txBuf, sendLength);
    if (st != NDIS_STATUS_SUCCESS) {
        wsprintfW(tm, L"TX: send FAIL len=%u", packetLength);
        LogMsg(tm);
        g_TxFail++;
        return NDIS_STATUS_FAILURE;
    }

    wsprintfW(tm, L"TX: send OK len=%u", packetLength);
    LogMsg(tm);
    g_TxOk++;
    return NDIS_STATUS_SUCCESS;
}

/* ============================================================================
 * OID support
 * ============================================================================ */
static const NDIS_OID g_SupportedOids[] = {
    OID_GEN_SUPPORTED_LIST,
    OID_GEN_HARDWARE_STATUS,
    OID_GEN_MEDIA_SUPPORTED,
    OID_GEN_MEDIA_IN_USE,
    OID_GEN_MAXIMUM_LOOKAHEAD,
    OID_GEN_MAXIMUM_FRAME_SIZE,
    OID_GEN_LINK_SPEED,
    OID_GEN_TRANSMIT_BUFFER_SPACE,
    OID_GEN_RECEIVE_BUFFER_SPACE,
    OID_GEN_TRANSMIT_BLOCK_SIZE,
    OID_GEN_RECEIVE_BLOCK_SIZE,
    OID_GEN_VENDOR_ID,
    OID_GEN_VENDOR_DESCRIPTION,
    OID_GEN_CURRENT_PACKET_FILTER,
    OID_GEN_CURRENT_LOOKAHEAD,
    OID_GEN_DRIVER_VERSION,
    OID_GEN_MAXIMUM_TOTAL_SIZE,
    OID_GEN_MAC_OPTIONS,
    OID_GEN_MEDIA_CONNECT_STATUS,
    OID_GEN_MAXIMUM_SEND_PACKETS,
    OID_GEN_VENDOR_DRIVER_VERSION,
    OID_GEN_XMIT_OK,
    OID_GEN_RCV_OK,
    OID_GEN_XMIT_ERROR,
    OID_GEN_RCV_ERROR,
    OID_GEN_RCV_NO_BUFFER,
    OID_802_3_PERMANENT_ADDRESS,
    OID_802_3_CURRENT_ADDRESS,
    OID_802_3_MULTICAST_LIST,
    OID_802_3_MAXIMUM_LIST_SIZE,
    OID_802_3_RCV_ERROR_ALIGNMENT,
    OID_802_3_XMIT_ONE_COLLISION,
    OID_802_3_XMIT_MORE_COLLISIONS,
    OID_PNP_CAPABILITIES,
    OID_PNP_QUERY_POWER,
    OID_PNP_SET_POWER
};

static const char g_VendorDesc[] = "Realtek RTL8152B USB Ethernet";

static NDIS_STATUS RtlQueryInformation(
    IN NDIS_HANDLE MiniportAdapterContext,
    IN NDIS_OID Oid,
    IN PVOID InformationBuffer,
    IN ULONG InformationBufferLength,
    OUT PULONG BytesWritten,
    OUT PULONG BytesNeeded)
{
    PRTL8152_ADAPTER A = (PRTL8152_ADAPTER)MiniportAdapterContext;
    NDIS_STATUS Status = NDIS_STATUS_SUCCESS;
    ULONG  GenUlong;
    USHORT GenUshort;
    PVOID  Src  = &GenUlong;
    ULONG  SrcLen = sizeof(ULONG);
    UCHAR  VendorId[4];
    NDIS_PNP_CAPABILITIES PnpCaps;

    *BytesWritten = 0;
    *BytesNeeded  = 0;

    switch (Oid) {
    case OID_GEN_SUPPORTED_LIST:
        Src = (PVOID)g_SupportedOids;
        SrcLen = sizeof(g_SupportedOids);
        break;
    case OID_GEN_HARDWARE_STATUS:
        GenUlong = NdisHardwareStatusReady;
        break;
    case OID_GEN_MEDIA_SUPPORTED:
    case OID_GEN_MEDIA_IN_USE:
        GenUlong = NdisMedium802_3;
        break;
    case OID_GEN_MAXIMUM_LOOKAHEAD:
    case OID_GEN_MAXIMUM_FRAME_SIZE:
        GenUlong = 1500;
        break;
    case OID_GEN_CURRENT_LOOKAHEAD:
        GenUlong = g_LookAhead;
        break;
    case OID_GEN_LINK_SPEED:
        GenUlong = (A && A->CurrentLinkSpeed) ? A->CurrentLinkSpeed : 1000000;
        break;
    case OID_GEN_TRANSMIT_BUFFER_SPACE:
        GenUlong = 1536;
        break;
    case OID_GEN_RECEIVE_BUFFER_SPACE:
        GenUlong = 16384;
        break;
    case OID_GEN_TRANSMIT_BLOCK_SIZE:
    case OID_GEN_RECEIVE_BLOCK_SIZE:
        GenUlong = 1536;
        break;
    case OID_GEN_VENDOR_ID:
        memcpy(VendorId, A->MacAddress, 3);
        VendorId[3] = 0;
        Src = VendorId;
        SrcLen = 4;
        break;
    case OID_GEN_VENDOR_DESCRIPTION:
        Src = (PVOID)g_VendorDesc;
        SrcLen = sizeof(g_VendorDesc);
        break;
    case OID_GEN_CURRENT_PACKET_FILTER:
        GenUlong = g_PacketFilter;
        break;
    case OID_GEN_DRIVER_VERSION:
        GenUshort = 0x0501;
        Src = &GenUshort;
        SrcLen = sizeof(USHORT);
        break;
    case OID_GEN_MAXIMUM_TOTAL_SIZE:
        GenUlong = 1514;
        break;
    case OID_GEN_MAC_OPTIONS:
        GenUlong = NDIS_MAC_OPTION_TRANSFERS_NOT_PEND
                 | NDIS_MAC_OPTION_COPY_LOOKAHEAD_DATA
                 | NDIS_MAC_OPTION_NO_LOOPBACK;
        break;
    case OID_GEN_MEDIA_CONNECT_STATUS:
        GenUlong = (A && A->bMediaConnected)
                 ? NdisMediaStateConnected
                 : NdisMediaStateDisconnected;
        break;
    case OID_GEN_MAXIMUM_SEND_PACKETS:
        GenUlong = 1;
        break;
    case OID_GEN_VENDOR_DRIVER_VERSION:
        GenUlong = 0x00010001;  /* 1.1 (Build 001) */
        break;
    case OID_GEN_XMIT_OK:
        GenUlong = g_TxOk;
        break;
    case OID_GEN_RCV_OK:
        GenUlong = g_RxPkts;
        break;
    case OID_GEN_XMIT_ERROR:
        GenUlong = g_TxFail;
        break;
    case OID_GEN_RCV_ERROR:
        GenUlong = g_RxErr;
        break;
    case OID_GEN_RCV_NO_BUFFER:
        GenUlong = 0;
        break;
    case OID_802_3_PERMANENT_ADDRESS:
    case OID_802_3_CURRENT_ADDRESS:
        Src = A->MacAddress;
        SrcLen = 6;
        break;
    case OID_802_3_MULTICAST_LIST:
        Src = NULL;
        SrcLen = 0;
        *BytesWritten = 0;
        return NDIS_STATUS_SUCCESS;
    case OID_802_3_MAXIMUM_LIST_SIZE:
        GenUlong = 32;
        break;
    case OID_802_3_RCV_ERROR_ALIGNMENT:
    case OID_802_3_XMIT_ONE_COLLISION:
    case OID_802_3_XMIT_MORE_COLLISIONS:
        GenUlong = 0;
        break;
    case OID_PNP_CAPABILITIES:
        memset(&PnpCaps, 0, sizeof(PnpCaps));
        PnpCaps.WakeUpCapabilities.MinMagicPacketWakeUp = NdisDeviceStateUnspecified;
        PnpCaps.WakeUpCapabilities.MinPatternWakeUp     = NdisDeviceStateUnspecified;
        PnpCaps.WakeUpCapabilities.MinLinkChangeWakeUp  = NdisDeviceStateUnspecified;
        Src = &PnpCaps;
        SrcLen = sizeof(PnpCaps);
        break;
    case OID_PNP_QUERY_POWER:
        return NDIS_STATUS_SUCCESS;
    default:
        Status = NDIS_STATUS_INVALID_OID;
        break;
    }

    if (Status == NDIS_STATUS_SUCCESS && Src != NULL) {
        if (InformationBufferLength < SrcLen) {
            *BytesNeeded = SrcLen;
            Status = NDIS_STATUS_INVALID_LENGTH;
        } else {
            memcpy(InformationBuffer, Src, SrcLen);
            *BytesWritten = SrcLen;
        }
    }
    return Status;
}

static NDIS_STATUS RtlSetInformation(
    IN NDIS_HANDLE MiniportAdapterContext,
    IN NDIS_OID Oid,
    IN PVOID InformationBuffer,
    IN ULONG InformationBufferLength,
    OUT PULONG BytesRead,
    OUT PULONG BytesNeeded)
{
    (void)MiniportAdapterContext;
    *BytesRead = 0;
    *BytesNeeded = 0;

    switch (Oid) {
    case OID_GEN_CURRENT_PACKET_FILTER:
        if (InformationBufferLength < sizeof(ULONG)) {
            *BytesNeeded = sizeof(ULONG);
            return NDIS_STATUS_INVALID_LENGTH;
        }
        memcpy(&g_PacketFilter, InformationBuffer, sizeof(ULONG));
        {
            WCHAR pm[64];
            wsprintfW(pm, L"OID: PacketFilter=%08X", g_PacketFilter);
            LogMsg(pm);
        }
        if (InterlockedCompareExchange(&g_bAttached, 1, 1) == 1) {
            ApplyRxPacketFilter(g_PacketFilter);
            if (g_PacketFilter != 0) {
                InterlockedExchange(&g_RxIndicateReady, 1);
                LogMsg(L"OID: RX indicate READY");
            }
        }
        *BytesRead = sizeof(ULONG);
        return NDIS_STATUS_SUCCESS;

    case OID_GEN_CURRENT_LOOKAHEAD:
        if (InformationBufferLength < sizeof(ULONG)) {
            *BytesNeeded = sizeof(ULONG);
            return NDIS_STATUS_INVALID_LENGTH;
        }
        memcpy(&g_LookAhead, InformationBuffer, sizeof(ULONG));
        *BytesRead = sizeof(ULONG);
        return NDIS_STATUS_SUCCESS;

    case OID_802_3_MULTICAST_LIST:
        *BytesRead = InformationBufferLength;
        return NDIS_STATUS_SUCCESS;

    case OID_PNP_SET_POWER:
        *BytesRead = sizeof(ULONG);
        return NDIS_STATUS_SUCCESS;

    default:
        return NDIS_STATUS_INVALID_OID;
    }
}
