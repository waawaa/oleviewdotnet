# Changes — Windows 11 24H2 (Build 26100) Compatibility Fix

## Summary

OleViewDotNet's COM process parser crashed with `NtApiDotNet.NtException` when enumerating
IPID entries on Windows 11 24H2 (build 26100). The root cause is that Microsoft changed the
layout of `tagIPIDEntry` in `combase.dll`, adding two new fields. This caused every field
from `cStrongRefs` onward to be read at the wrong offset, making `pOXIDEntry` read the value
of `pStub` instead — a COM stub pointer, not an OXID entry address.

---

## Root Cause: `tagIPIDEntry` Struct Layout Change in combase.dll

### What changed in Windows 11 24H2 (build 26100)

Microsoft added **two new fields** to the native `tagIPIDEntry` struct in `combase.dll`:

| Offset | Field             | Size  | Status       |
|--------|-------------------|-------|--------------|
| +0x00  | pNextIPID         | 8     | Unchanged    |
| +0x08  | dwFlags           | 4     | Unchanged    |
| +0x0C  | **flags**         | **4** | **NEW**      |
| +0x10  | cStrongRefs       | 4     | Shifted +4   |
| +0x14  | cWeakRefs         | 4     | Shifted +4   |
| +0x18  | cPrivateRefs      | 4     | Shifted +4   |
| +0x1C  | *(padding)*       | 4     | Shifted +4   |
| +0x20  | pv                | 8     | Shifted +8   |
| +0x28  | pStub             | 8     | Shifted +8   |
| +0x30  | pOXIDEntry        | 8     | Shifted +8   |
| +0x38  | ipid              | 16    | Shifted +8   |
| +0x48  | iid               | 16    | Shifted +8   |
| +0x58  | pChnl             | 8     | Shifted +8   |
| +0x60  | pIRCEntry         | 8     | Shifted +8   |
| +0x68  | **pInterfaceName**| **8** | **NEW**      |
| +0x70  | pOIDFLink         | 8     | Shifted +16  |
| +0x78  | pOIDBLink         | 8     | Shifted +16  |

**Old size**: 0x70 (112 bytes, 64-bit)
**New size**: 0x80 (128 bytes, 64-bit)

### New fields

1. **`flags`** at +0x0C — `std::atomic<unsigned long>` (4 bytes)
   Inserted between `dwFlags` and `cStrongRefs`. This is a separate atomic flags field
   distinct from the existing `dwFlags`.

2. **`pInterfaceName`** at +0x68 — `Ptr64 HSTRING__` (8 bytes)
   Inserted between `pIRCEntry` and `pOIDFLink`. Stores an HSTRING pointer to the
   interface name, likely for WinRT/UWP interface tracking and diagnostics.

### How this broke OleViewDotNet

The old `IPIDEntryNative` C# struct mapped fields at their pre-26100 offsets. With the new
layout, every field from `cStrongRefs` onward was read from the wrong memory position:

- **`pOXIDEntry`** (code offset 0x28) read the native **`pStub`** field (native offset 0x28)
  instead of the real `pOXIDEntry` (native offset 0x30).
- The misread `pStub` value pointed to a COM stub (IUnknown vtable), not an `OXIDEntry`.
- When `ReadStruct<OXIDEntryNative2004>()` tried to interpret the stub address as an OXID
  entry, it read freed/reused memory containing Unicode string data
  (e.g., `"Shell"`, `"mshelp://help/?id=escalation"`).
- The `_psa` field within the OXID entry thus contained garbage like `0x006500680073006d`
  (Unicode `"mshe"`), which was passed to `COMDualStringArray(IntPtr, NtProcess)`.
- `NtVirtualMemory.ReadMemory<ushort>()` attempted to dereference this garbage pointer in
  the target process, throwing `NtException`.

**Proof from live debugging (WinDbg/CDB attached to target process):**
```
Code reads:  pChnl = 0x46000000000000C0
Expected:    This is the last 8 bytes of IID_IUnknown {00000000-0000-0000-C000-000000000046}
             confirming the 8-byte offset shift (code offset 0x50 reads native offset 0x50
             which is the tail of iid, not pChnl).
```

### Verification method

The struct layout was confirmed by dumping the actual types from `combase.dll` private symbols
on the target process using WinDbg:

```
0:000> dt combase!tagIPIDEntry
   +0x000 pNextIPID        : Ptr64 tagIPIDEntry
   +0x008 dwFlags          : Uint4B
   +0x00c flags            : std::atomic<unsigned long>    <-- NEW
   +0x010 cStrongRefs      : Uint4B
   +0x014 cWeakRefs        : Uint4B
   +0x018 cPrivateRefs     : Uint4B
   +0x020 pv               : Ptr64 Void
   +0x028 pStub            : Ptr64 IUnknown
   +0x030 pOXIDEntry       : Ptr64 OXIDEntry
   +0x038 ipid             : _GUID
   +0x048 iid              : _GUID
   +0x058 pChnl            : Ptr64 CCtxComChnl
   +0x060 pIRCEntry        : Ptr64 IRCEntry
   +0x068 pInterfaceName   : Ptr64 HSTRING__               <-- NEW
   +0x070 pOIDFLink        : Ptr64 tagIPIDEntry
   +0x078 pOIDBLink        : Ptr64 tagIPIDEntry
```

The `__MIDL_ILocalObjectExporter_0007` (OXID info) struct was also verified and found to be
**unchanged** — the `OXIDInfoNative2004` / `OXIDEntryNative2004` layouts remain correct.

---

## Files Changed

### New Files

#### `OleViewDotNet/Processes/Types/IPIDEntryNative26100.cs`
New 64-bit IPID entry struct matching the Windows 26100+ `tagIPIDEntry` layout.
Adds the `flags` (uint) and `pInterfaceName` (IntPtr) fields at the correct offsets.
Total marshaled size: 0x80 (128 bytes).

#### `OleViewDotNet/Processes/Types/IPIDEntryNative26100_32.cs`
New 32-bit IPID entry struct matching the Windows 26100+ `tagIPIDEntry` layout for
WoW64 processes. Adds the same two new fields with 32-bit pointer sizes.

### Modified Files

#### `OleViewDotNet/Processes/COMProcessParser.cs`
- **`ParseIPIDEntries` (non-generic overload)**: Now reads the `CPageAllocator.EntrySize`
  from the target process *before* selecting the IPID struct type. If `EntrySize` is large
  enough for the new struct (`>= Marshal.SizeOf(IPIDEntryNative26100)`), the 26100 variant
  is used; otherwise, the original struct is used. This makes the detection automatic and
  forward-compatible — no hard-coded Windows version check needed.
- **`ParseIPIDEntries<T>` (generic overload)**: Added per-entry `try/catch (NtException)`
  around `new COMIPIDEntry(...)` so that one malformed or inaccessible IPID entry does not
  abort the entire enumeration.

#### `OleViewDotNet/Marshaling/COMDualStringArray.cs`
Defense-in-depth hardening of the `COMDualStringArray(IntPtr, NtProcess)` constructor:
- Added `MaxDualStringArrayEntries = 4096` upper bound on `num_entries` to reject obviously
  corrupt values.
- Added length validation on the `ReadMemory` return to detect short reads.
- Added bounds check on `sec_offset` before seeking in `ReadEntries` to prevent
  `ArgumentOutOfRangeException`.

#### `OleViewDotNet/OleViewDotNet.csproj`
- Added `<GenerateResourceUsePreserializedResources>true</GenerateResourceUsePreserializedResources>`
  to fix MSB3823/MSB3822 build errors with newer .NET SDK versions.
- Added `<PackageReference Include="System.Resources.Extensions" Version="8.0.0" />`.

#### `third_party/NtApiDotNet/NtApiDotNet.csproj`
- Same `GenerateResourceUsePreserializedResources` and `System.Resources.Extensions` fixes
  for the NtApiDotNet project.

---

## Build & Test

```
dotnet build OleViewDotNet.sln --configuration Release
```

Verified: 0 warnings, 0 errors. Tested on Windows 11 24H2 (build 26100) x64.
