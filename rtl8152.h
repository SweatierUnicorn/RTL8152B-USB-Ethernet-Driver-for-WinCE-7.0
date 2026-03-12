/* ============================================================================
 * rtl8152.h  -  Realtek RTL8152B USB-to-Fast-Ethernet NDIS Miniport
 *               Register map, constants, adapter context.
 *               Target: WinCE 7.0 / ARM (Panasonic Strada CN-F1X10BD)
 *
 * Author      : SweatierUnicorn
 * GitHub      : https://github.com/SweatierUnicorn
 * Distribution: Provided free of charge by the author.
 *
 * Reference: Linux r8152.c (Realtek)
 * C89 only.
 * ============================================================================ */
#ifndef RTL8152_H
#define RTL8152_H

#include <windows.h>
#include <ndis.h>
#include <usbdi.h>

/* ============================
 * USB Control Transfer
 * ============================ */
#define RTL8152_REQT_READ       0xC0    /* bmRequestType: vendor, device-to-host */
#define RTL8152_REQT_WRITE      0x40    /* bmRequestType: vendor, host-to-device */
#define RTL8152_REQ_GET_REGS    0x05    /* bRequest for read */
#define RTL8152_REQ_SET_REGS    0x05    /* bRequest for write */

/* MCU type = wValue for control transfer */
#define MCU_TYPE_PLA            0x0100  /* PLA register space */
#define MCU_TYPE_USB            0x0000  /* USB register space */

/* Byte-enable masks for USB vendor writes.
 * These are the raw Realtek/Linux values passed into generic_ocp_write():
 *   DWORD = 0xFF, WORD = 0x33, BYTE = 0x11
 * For upper byte/word accesses they are shifted left by byte offset.
 * Example: upper word => 0x33 << 2 = 0xCC. */
#define BYTE_EN_DWORD           0xFF
#define BYTE_EN_WORD            0x33
#define BYTE_EN_BYTE            0x11
#define BYTE_EN_SIX_BYTES       0x3F    /* for MAC address read */
#define BYTE_EN_START_MASK      0x0F
#define BYTE_EN_END_MASK        0xF0

/* ============================
 * PLA Registers
 * ============================ */

/* MAC Address (6 bytes) */
#define PLA_IDR                 0xC000

/* RX Control */
#define PLA_RCR                 0xC010
#define   RCR_AAP               0x00000001  /* accept all physical (promisc) */
#define   RCR_APM               0x00000002  /* accept physical match (unicast) */
#define   RCR_AM                0x00000004  /* accept multicast */
#define   RCR_AB                0x00000008  /* accept broadcast */
#define   RCR_ACPT_ALL          (RCR_AAP | RCR_APM | RCR_AM | RCR_AB)

/* RX Max Size */
#define PLA_RMS                 0xC016

/* Multicast Address filter */
#define PLA_MAR                 0xCD00

/* RX FIFO Control */
#define PLA_RXFIFO_CTRL0        0xC0A0
#define PLA_RXFIFO_FULL         0xC0A2
#define PLA_RXFIFO_CTRL1        0xC0A4
#define PLA_RX_FIFO_FULL        0xC0A6
#define PLA_RXFIFO_CTRL2        0xC0A8
#define PLA_RX_FIFO_EMPTY       0xC0AA

/* TX FIFO Control */
#define PLA_TXFIFO_CTRL         0xE0B0
#define PLA_TX_FIFO_FULL        0xE0B2
#define PLA_TX_FIFO_EMPTY       0xE0B4

/* TX Control */
#define PLA_TCR0                0xE610  /* also used for version detection */
#define PLA_TCR1                0xE612
#define   TCR0_TX_EMPTY         0x0800
#define   TCR0_AUTO_FIFO        0x0080

/* Chip Control */
#define PLA_CR                  0xE813
#define   CR_RST                0x10    /* software reset */
#define   CR_RE                 0x08    /* RX enable */
#define   CR_TE                 0x04    /* TX enable */

/* Config Write Enable */
#define PLA_CRWECR              0xE81C
#define   CRWECR_NORAML         0x00    /* normal mode (writes disabled) */
#define   CRWECR_CONFIG         0xC0    /* config mode (writes enabled) */

/* Physical Link Status */
#define PLA_PHYSTATUS           0xE908
#define   LINK_STATUS           0x02    /* bit 1: 1=link up */

/* SEN (Software-Enable Network) */
#define PLA_SFF_STS_7           0xE8DE
#define   RE_INIT_LL            0x8000
#define   MCU_BORW_EN           0x4000

/* Chip Version */
#define VERSION_MASK            0x7CF0

/* Decoded versions */
#define RTL_VER_01              0       /* RTL8152 rev A */
#define RTL_VER_02              1       /* RTL8152 rev B */
#define RTL_VER_03              2       /* RTL8153 rev A */
#define RTL_VER_04              3       /* RTL8153 rev B */
#define RTL_VER_05              4       /* RTL8153 rev C */

/* PLA_MISC_1 */
#define PLA_MISC_1              0xE85A
#define   RXDY_GATED_EN         0x0008

/* PLA_FMC - Packet Filter MCU control */
#define PLA_FMC                 0xC0B4
#define   FMC_FCR_MCU_EN        0x0001

/* PLA_LED_FEATURE */
#define PLA_LED_FEATURE         0xDD92
#define   LED_MODE_MASK         0x0700

/* PLA_PHY_PWR - PHY power control */
#define PLA_PHY_PWR             0xE84C
#define   TX_10M_IDLE_EN        0x0080
#define   PFM_PWM_SWITCH        0x0040

/* PLA_MAC_PWR_CTRL - MCU clock gating */
#define PLA_MAC_PWR_CTRL        0xE0C0
#define   D3_CLK_GATED_EN       0x00004000
#define   MCU_CLK_RATIO         0x07010F07
#define   MCU_CLK_RATIO_MASK    0x0F0F0F0F

/* PLA_GPHY_INTR_IMR - GPHY interrupt mask */
#define PLA_GPHY_INTR_IMR       0xE022
#define   GPHY_STS_MSK          0x0001
#define   SPEED_DOWN_MSK        0x0002
#define   SPDWN_RXDV_MSK        0x0004
#define   SPDWN_LINKCHG_MSK     0x0008

#define PLA_OCP_GPHY_BASE       0xE86C
#define PLA_TALLYCNT            0xE890
#define PLA_TEREDO_CFG          0xE884
#define PLA_EXTRA_STATUS        0xE80A

#define PLA_CFG_WOL             0xE8BC
#define   MAGIC_EN              0x0001

#define PLA_CPCR                0xE854
#define   CPCR_RX_VLAN          0x0040

#define PLA_BOOT_CTRL           0xE004
#define   AUTOLOAD_DONE         0x0002

/* ============================
 * USB-specific registers
 * ============================ */
#define USB_USB2PHY             0xB41E
#define USB_SSPHYLINK2          0xB428
#define USB_USB_CTRL            0xD406
#define   RX_AGG_DISABLE        0x0010
#define   RX_ZERO_EN            0x0080
#define USB_USB_TIMER           0xD428
#define USB_RX_BUF_TH          0xD40C
#define USB_TX_AGG              0xD40A
#define USB_RX_EARLY_AGG        0xD40E
#define USB_UPS_CTRL            0xD800  /* power cut control (RTL8152) */
#define   POWER_CUT             0x0100
#define USB_PM_CTRL_STATUS      0xD432  /* RTL8153A only - RESUME_INDICATE */
#define   RESUME_INDICATE       0x0001
#define USB_MSC_TIMER           0xCBFC
#define USB_TX_DMA              0xD434
#define USB_TOLERANCE           0xD490
#define USB_BMU_RESET           0xD4B0
#define   BMU_RESET_EP_IN       0x01
#define   BMU_RESET_EP_OUT      0x02
#define USB_BMU_CONFIG          0xD4B4

/* ============================
 * OCP/PHY Registers (via indirect MII)
 * ============================ */
#define OCP_BASE_MII            0xA400
#define OCP_PHY_STATUS          0xA420
#define OCP_ALDPS_CONFIG        0x2010
#define   ENPWRSAVE             0x8000
#define   ENPDNPS               0x0200
#define   LINKENA               0x0100
#define   DIS_SDSAVE            0x0010  /* bit 4, NOT same as LINKENA */

/* MII PHY registers (standard) */
#define MII_BMCR                0x00
#define   BMCR_RESET            0x8000
#define   BMCR_ANENABLE         0x1000
#define   BMCR_ANRESTART        0x0200
#define   BMCR_SPEED100         0x2000
#define   BMCR_FULLDPLX         0x0100
#define   BMCR_PDOWN            0x0800  /* power-down bit */

#define MII_BMSR                0x01
#define   BMSR_LSTATUS          0x0004
#define   BMSR_ANEGCOMPLETE     0x0020

#define MII_ANAR                0x04
#define   ADVERTISE_10HALF      0x0020
#define   ADVERTISE_10FULL      0x0040
#define   ADVERTISE_100HALF     0x0080
#define   ADVERTISE_100FULL     0x0100
#define   ADVERTISE_CSMA        0x0001
#define   ADVERTISE_PAUSE_CAP   0x0400  /* flow control pause */
#define   ADVERTISE_PAUSE_ASYM  0x0800  /* asymmetric pause */
#define   ADVERTISE_ALL         (ADVERTISE_10HALF | ADVERTISE_10FULL | \
                                 ADVERTISE_100HALF | ADVERTISE_100FULL)

#define MII_ANLPAR              0x05

/* ============================
 * RX/TX Descriptor Structures
 * ============================ */

/* RX descriptor: 24 bytes prepended to each packet in bulk IN */
#pragma pack(push, 1)
typedef struct _RTL_RX_DESC {
    DWORD opts1;    /* bit 15: CRC ok, bits 14-0: packet length */
    DWORD opts2;
    DWORD opts3;
    DWORD opts4;
    DWORD opts5;
    DWORD opts6;
} RTL_RX_DESC, *PRTL_RX_DESC;
#pragma pack(pop)
#define RTL_RX_DESC_SIZE        24
#define RTL_RX_LEN_MASK         0x7FFF  /* bits 14:0 of opts1 */

/* TX descriptor: 8 bytes prepended to each packet in bulk OUT */
#pragma pack(push, 1)
typedef struct _RTL_TX_DESC {
    DWORD opts1;    /* bit 31: TX_FS, bit 30: TX_LS, bits 17:0: length */
    DWORD opts2;
} RTL_TX_DESC, *PRTL_TX_DESC;
#pragma pack(pop)
#define RTL_TX_DESC_SIZE        8
#define TX_FS                   0x80000000  /* first segment */
#define TX_LS                   0x40000000  /* last segment */
#define TX_LEN_MAX              0x0003FFFFUL

/* RTL8152B v1 RX descriptor */
#define RTL_RX_DESC_SIZE        24
#define RX_LEN_MASK             0x00007FFFUL
#define RD_CRC                  0x00008000UL
#define RTL_RX_ALIGN            8
#define ETH_HEADER_SIZE         14
#define ETH_FCS_LEN             4

/* ============================
 * Buffer sizes
 * ============================ */
#define RTL_AGG_BUF_SIZE        16384   /* 16KB aggregation buffer */
#define RTL_RX_BUF_SIZE         (48 * 1024)
#define RTL_MAX_ETH_SIZE        1536    /* max ethernet frame */
#define RTL_MIN_ETH_PAYLOAD     60

/* Basic TX init values from Linux r8152.c */
#define TX_AGG_MAX_THRESHOLD    0x03
#define RX_THR_HIGH             0x7A120180UL
#define TEST_MODE_DISABLE       0x00000001UL
#define TX_SIZE_ADJUST1         0x00000100UL
#define RTL8152_RMS             1518

/* ============================
 * Adapter Context
 * ============================ */
typedef struct _RTL8152_ADAPTER {
    NDIS_HANDLE             MiniportAdapterHandle;
    USB_HANDLE              UsbDevice;
    LPCUSB_FUNCS            UsbFuncs;

    UCHAR                   MacAddress[6];
    BOOL                    bHasMac;

    volatile BOOL           bMediaConnected;
    ULONG                   CurrentLinkSpeed;   /* 100 bps units */

    USB_PIPE                BulkInPipe;
    USB_PIPE                BulkOutPipe;
    USB_PIPE                InterruptPipe;

    HANDLE                  hRxThread;
    volatile BOOL           bExitRxThread;
    volatile BOOL           bUsbGone;

    UCHAR                   ChipVersion;        /* RTL_VER_01..05 */
} RTL8152_ADAPTER, *PRTL8152_ADAPTER;

#endif /* RTL8152_H */
