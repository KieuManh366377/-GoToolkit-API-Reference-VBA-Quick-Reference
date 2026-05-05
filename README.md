# GoToolkit API Reference — VBA Quick Reference

> **DLL:** `GoToolkit.dll` | **Version:** 1.0.0 | **Tổng:** 175 hàm | **Ngày:** 05/05/2026

---

## Lưu ý

| Kiểu trả về | Ý nghĩa |
|---|---|
| `LongPtr` | Con trỏ chuỗi UTF-8 — cần `FreeString` sau dùng |
| `Long` | Số nguyên hoặc mã lỗi (0 = lỗi / 1 = OK) |
| `LongLong` | Số nguyên 64-bit |
| `Double` | Số thực |

---

## Mục lục

- [bcrypt](#bcrypt)
- [compression](#compression)
- [converter](#converter)
- [crypto](#crypto)
- [datetime](#datetime)
- [encodeutils](#encodeutils)
- [filecompression](#filecompression)
- [fileio](#fileio)
- [hash](#hash)
- [http](#http)
- [jwt](#jwt)
- [stringutils](#stringutils)
- [validation](#validation)
- [zipfile](#zipfile)

---

## bcrypt

| Hàm | Khai báo VBA |
|---|---|
| `BcryptHashPassword`<br><sub>==========================================</sub> | `Declare PtrSafe Function BcryptHashPassword Lib "GoToolkit.dll" (ByVal password As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
PASSWORD HASH / VERIFY (mô phỏng BCrypt)
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `password` | LongPtr |

**Trả về:** LongPtr

| `BcryptVerifyPassword` | `Declare PtrSafe Function BcryptVerifyPassword Lib "GoToolkit.dll" (ByVal password As LongPtr, ByVal hashed As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `password` | LongPtr |
| `hashed` | LongPtr |

**Trả về:** Long

| `BcryptEncryptString`<br><sub>==========================================</sub> | `Declare PtrSafe Function BcryptEncryptString Lib "GoToolkit.dll" (ByVal plaintext As LongPtr, ByVal password As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
MÃ HOÁ CHUỖI AES-256-GCM
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `plaintext` | LongPtr |
| `password` | LongPtr |

**Trả về:** LongPtr

| `BcryptDecryptString` | `Declare PtrSafe Function BcryptDecryptString Lib "GoToolkit.dll" (ByVal encrypted As LongPtr, ByVal password As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `encrypted` | LongPtr |
| `password` | LongPtr |

**Trả về:** LongPtr

| `BcryptEncryptFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function BcryptEncryptFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr, ByVal password As LongPtr) As Long` |

**Mô tả:** ==========================================
MÃ HOÁ FILE AES-256-GCM
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcPath` | LongPtr |
| `dstPath` | LongPtr |
| `password` | LongPtr |

**Trả về:** Long

| `BcryptDecryptFile` | `Declare PtrSafe Function BcryptDecryptFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr, ByVal password As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcPath` | LongPtr |
| `dstPath` | LongPtr |
| `password` | LongPtr |

**Trả về:** Long

| `BcryptSignString`<br><sub>==========================================</sub> | `Declare PtrSafe Function BcryptSignString Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal secretKey As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
HMAC - KÝ VÀ XÁC THỰC FILE/CHUỖI
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `secretKey` | LongPtr |

**Trả về:** LongPtr

| `BcryptVerifyString` | `Declare PtrSafe Function BcryptVerifyString Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal signature As LongPtr, ByVal secretKey As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `signature` | LongPtr |
| `secretKey` | LongPtr |

**Trả về:** Long

| `BcryptSignFile` | `Declare PtrSafe Function BcryptSignFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal secretKey As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |
| `secretKey` | LongPtr |

**Trả về:** LongPtr

| `BcryptVerifyFile` | `Declare PtrSafe Function BcryptVerifyFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal signature As LongPtr, ByVal secretKey As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |
| `signature` | LongPtr |
| `secretKey` | LongPtr |

**Trả về:** Long

| `BcryptFreeString` | `Declare PtrSafe Sub BcryptFreeString Lib "GoToolkit.dll" (ByVal s As LongPtr)` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Không có (Sub)

---

## compression

| Hàm | Khai báo VBA |
|---|---|
| `GzipCompress`<br><sub>==========================================</sub> | `Declare PtrSafe Function GzipCompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
GZIP
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `GzipDecompress` | `Declare PtrSafe Function GzipDecompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `ZlibCompress`<br><sub>==========================================</sub> | `Declare PtrSafe Function ZlibCompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
ZLIB
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `ZlibDecompress` | `Declare PtrSafe Function ZlibDecompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `CompressRatio`<br><sub>==========================================</sub> | `Declare PtrSafe Function CompressRatio Lib "GoToolkit.dll" (ByVal original As LongPtr, ByVal compressed As LongPtr) As Long` |

**Mô tả:** ==========================================
Tính tỉ lệ nén
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `original` | LongPtr |
| `compressed` | LongPtr |

**Trả về:** Long

| `FreeString` | `Declare PtrSafe Sub FreeString Lib "GoToolkit.dll" (ByVal s As LongPtr)` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Không có (Sub)

---

## converter

| Hàm | Khai báo VBA |
|---|---|
| `GT_Mem_Copy`<br><sub>GT_Mem_Copy: Thay the CopyMemory cua kernel32</sub> | `Declare PtrSafe Sub GT_Mem_Copy Lib "GoToolkit.dll" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)` |

**Mô tả:** GT_Mem_Copy: Thay the CopyMemory cua kernel32
Dung memmove -> an toan khi vung nho chong lan

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `destination` | LongPtr |
| `source` | LongPtr |
| `length` | LongPtr |

**Trả về:** Không có (Sub)

| `GT_Mem_PtrToStringUTF8`<br><sub>GT_Mem_PtrToStringUTF8: Sao chep chuoi null-terminated tu...</sub> | `Declare PtrSafe Function GT_Mem_PtrToStringUTF8 Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |

**Mô tả:** GT_Mem_PtrToStringUTF8: Sao chep chuoi null-terminated tu dia chi bo nho
- Tra 0 neu ptr=0 hoac chuoi vuot 64MB
- Can goi GT_Mem_Free(ket_qua) sau khi dung xong

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |

**Trả về:** LongPtr

| `GT_Mem_Free`<br><sub>GT_Mem_Free: Giai phong bo nho tu GT_Mem_PtrToStringUTF8 ...</sub> | `Declare PtrSafe Sub GT_Mem_Free Lib "GoToolkit.dll" (ByVal ptr As LongPtr)` |

**Mô tả:** GT_Mem_Free: Giai phong bo nho tu GT_Mem_PtrToStringUTF8 / GT_PtrToString

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |

**Trả về:** Không có (Sub)

| `GT_Mem_PtrReadByte`<br><sub>GT_Mem_PtrReadByte: Doc 1 byte tai dia chi chi dinh</sub> | `Declare PtrSafe Function GT_Mem_PtrReadByte Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |

**Mô tả:** GT_Mem_PtrReadByte: Doc 1 byte tai dia chi chi dinh

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |

**Trả về:** LongPtr

| `GT_Mem_PtrReadInt32`<br><sub>GT_Mem_PtrReadInt32: Doc 4 byte (Long trong VBA)</sub> | `Declare PtrSafe Function GT_Mem_PtrReadInt32 Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As Long` |

**Mô tả:** GT_Mem_PtrReadInt32: Doc 4 byte (Long trong VBA)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |

**Trả về:** Long

| `GT_Mem_PtrReadInt64`<br><sub>GT_Mem_PtrReadInt64: Doc 8 byte (LongLong trong VBA 64-bit)</sub> | `Declare PtrSafe Function GT_Mem_PtrReadInt64 Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |

**Mô tả:** GT_Mem_PtrReadInt64: Doc 8 byte (LongLong trong VBA 64-bit)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |

**Trả về:** LongPtr

| `GT_Mem_PtrReadDouble`<br><sub>GT_Mem_PtrReadDouble: Doc 8 byte (Double trong VBA)</sub> | `Declare PtrSafe Function GT_Mem_PtrReadDouble Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As Double` |

**Mô tả:** GT_Mem_PtrReadDouble: Doc 8 byte (Double trong VBA)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |

**Trả về:** Double

| `GT_Mem_Zero`<br><sub>GT_Mem_Zero: Xoa trang vung nho (huu ich cho du lieu nhay...</sub> | `Declare PtrSafe Sub GT_Mem_Zero Lib "GoToolkit.dll" (ByVal ptr As LongPtr, ByVal length As LongPtr)` |

**Mô tả:** GT_Mem_Zero: Xoa trang vung nho (huu ich cho du lieu nhay cam: password, key)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |
| `length` | LongPtr |

**Trả về:** Không có (Sub)

| `GT_PtrToString`<br><sub>GT_PtrToString: Ten ngan hon cua GT_Mem_PtrToStringUTF8</sub> | `Declare PtrSafe Function GT_PtrToString Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |

**Mô tả:** GT_PtrToString: Ten ngan hon cua GT_Mem_PtrToStringUTF8
Tra ve LongPtr tro toi ban sao chuoi, can goi GT_Mem_Free sau

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ptr` | LongPtr |

**Trả về:** LongPtr

---

## crypto

| Hàm | Khai báo VBA |
|---|---|
| `SetKey` | `Declare PtrSafe Function SetKey Lib "GoToolkit.dll" (ByVal key As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `key` | LongPtr |

**Trả về:** Long

| `Encrypt` | `Declare PtrSafe Function Encrypt Lib "GoToolkit.dll" (ByVal plaintext As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `plaintext` | LongPtr |

**Trả về:** LongPtr

| `Decrypt` | `Declare PtrSafe Function Decrypt Lib "GoToolkit.dll" (ByVal ciphertext As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `ciphertext` | LongPtr |

**Trả về:** LongPtr

---

## datetime

| Hàm | Khai báo VBA |
|---|---|
| `DateTimeNow`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeNow Lib "GoToolkit.dll" As LongPtr` |

**Mô tả:** ==========================================
THỜI GIAN HIỆN TẠI
==========================================

**Trả về:** LongPtr

| `DateNow` | `Declare PtrSafe Function DateNow Lib "GoToolkit.dll" As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Trả về:** LongPtr

| `TimeNow` | `Declare PtrSafe Function TimeNow Lib "GoToolkit.dll" As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Trả về:** LongPtr

| `DateTimeUTC` | `Declare PtrSafe Function DateTimeUTC Lib "GoToolkit.dll" As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Trả về:** LongPtr

| `UnixTimestamp` | `Declare PtrSafe Function UnixTimestamp Lib "GoToolkit.dll" As LongLong` |

**Mô tả:** _Chưa có mô tả._

**Trả về:** LongLong

| `UnixTimestampMs` | `Declare PtrSafe Function UnixTimestampMs Lib "GoToolkit.dll" As LongLong` |

**Mô tả:** _Chưa có mô tả._

**Trả về:** LongLong

| `DateTimeFormat`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeFormat Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal fmtIn As LongPtr, ByVal fmtOut As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
ĐỊNH DẠNG
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |
| `fmtIn` | LongPtr |
| `fmtOut` | LongPtr |

**Trả về:** LongPtr

| `UnixToDateTime` | `Declare PtrSafe Function UnixToDateTime Lib "GoToolkit.dll" (ByVal unix As LongLong) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `unix` | LongLong |

**Trả về:** LongPtr

| `DateTimeToUnix` | `Declare PtrSafe Function DateTimeToUnix Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As LongLong` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** LongLong

| `DateTimeAddDays`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeAddDays Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nDays As Long) As LongPtr` |

**Mô tả:** ==========================================
TÍNH TOÁN
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |
| `nDays` | Long |

**Trả về:** LongPtr

| `DateTimeAddMonths` | `Declare PtrSafe Function DateTimeAddMonths Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nMonths As Long) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |
| `nMonths` | Long |

**Trả về:** LongPtr

| `DateTimeAddYears` | `Declare PtrSafe Function DateTimeAddYears Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nYears As Long) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |
| `nYears` | Long |

**Trả về:** LongPtr

| `DateTimeAddHours` | `Declare PtrSafe Function DateTimeAddHours Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nHours As Long) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |
| `nHours` | Long |

**Trả về:** LongPtr

| `DateTimeAddMinutes` | `Declare PtrSafe Function DateTimeAddMinutes Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nMins As Long) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |
| `nMins` | Long |

**Trả về:** LongPtr

| `DateTimeDiffDays`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeDiffDays Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |

**Mô tả:** ==========================================
SO SÁNH / KHOẢNG CÁCH
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr1` | LongPtr |
| `dtStr2` | LongPtr |

**Trả về:** Long

| `DateTimeDiffHours` | `Declare PtrSafe Function DateTimeDiffHours Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr1` | LongPtr |
| `dtStr2` | LongPtr |

**Trả về:** Long

| `DateTimeDiffMinutes` | `Declare PtrSafe Function DateTimeDiffMinutes Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr1` | LongPtr |
| `dtStr2` | LongPtr |

**Trả về:** Long

| `DateTimeDiffSeconds` | `Declare PtrSafe Function DateTimeDiffSeconds Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr1` | LongPtr |
| `dtStr2` | LongPtr |

**Trả về:** Long

| `DateTimeIsBefore` | `Declare PtrSafe Function DateTimeIsBefore Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr1` | LongPtr |
| `dtStr2` | LongPtr |

**Trả về:** Long

| `DateTimeIsAfter` | `Declare PtrSafe Function DateTimeIsAfter Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr1` | LongPtr |
| `dtStr2` | LongPtr |

**Trả về:** Long

| `DateGetYear`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateGetYear Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |

**Mô tả:** ==========================================
THÔNG TIN NGÀY
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** Long

| `DateGetMonth` | `Declare PtrSafe Function DateGetMonth Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** Long

| `DateGetDay` | `Declare PtrSafe Function DateGetDay Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** Long

| `DateGetWeekday` | `Declare PtrSafe Function DateGetWeekday Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** LongPtr

| `DateGetWeekNumber` | `Declare PtrSafe Function DateGetWeekNumber Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** Long

| `DateGetDayOfYear` | `Declare PtrSafe Function DateGetDayOfYear Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** Long

| `DateIsLeapYear` | `Declare PtrSafe Function DateIsLeapYear Lib "GoToolkit.dll" (ByVal nYear As Long) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `nYear` | Long |

**Trả về:** Long

| `DateGetDaysInMonth` | `Declare PtrSafe Function DateGetDaysInMonth Lib "GoToolkit.dll" (ByVal nYear As Long, ByVal nMonth As Long) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `nYear` | Long |
| `nMonth` | Long |

**Trả về:** Long

| `DateGetQuarter` | `Declare PtrSafe Function DateGetQuarter Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |

**Trả về:** Long

| `DateTimeConvertTZ`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeConvertTZ Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal fromTZ As LongPtr, ByVal toTZ As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
TIMEZONE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `dtStr` | LongPtr |
| `fromTZ` | LongPtr |
| `toTZ` | LongPtr |

**Trả về:** LongPtr

---

## encodeutils

| Hàm | Khai báo VBA |
|---|---|
| `Base64Encode`<br><sub>==========================================</sub> | `Declare PtrSafe Function Base64Encode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
BASE64
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `Base64Decode` | `Declare PtrSafe Function Base64Decode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `Base64EncodeURL` | `Declare PtrSafe Function Base64EncodeURL Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `Base64DecodeURL` | `Declare PtrSafe Function Base64DecodeURL Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `Base64IsValid` | `Declare PtrSafe Function Base64IsValid Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `HexEncode`<br><sub>==========================================</sub> | `Declare PtrSafe Function HexEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
HEX
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `HexDecode` | `Declare PtrSafe Function HexDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `HexIsValid` | `Declare PtrSafe Function HexIsValid Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `HexToDecimal` | `Declare PtrSafe Function HexToDecimal Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongLong` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongLong

| `DecimalToHex` | `Declare PtrSafe Function DecimalToHex Lib "GoToolkit.dll" (ByVal n As LongLong) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `n` | LongLong |

**Trả về:** LongPtr

| `DecimalToBinary` | `Declare PtrSafe Function DecimalToBinary Lib "GoToolkit.dll" (ByVal n As LongLong) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `n` | LongLong |

**Trả về:** LongPtr

| `BinaryToDecimal` | `Declare PtrSafe Function BinaryToDecimal Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongLong` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongLong

| `DecimalToOctal` | `Declare PtrSafe Function DecimalToOctal Lib "GoToolkit.dll" (ByVal n As LongLong) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `n` | LongLong |

**Trả về:** LongPtr

| `OctalToDecimal` | `Declare PtrSafe Function OctalToDecimal Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongLong` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongLong

| `URLEncode`<br><sub>==========================================</sub> | `Declare PtrSafe Function URLEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
URL ENCODE / DECODE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `URLDecode` | `Declare PtrSafe Function URLDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `URLPathEncode` | `Declare PtrSafe Function URLPathEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `URLPathDecode` | `Declare PtrSafe Function URLPathDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `HTMLEncode`<br><sub>==========================================</sub> | `Declare PtrSafe Function HTMLEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
HTML ENCODE / DECODE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `HTMLDecode` | `Declare PtrSafe Function HTMLDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `UTF8IsValid`<br><sub>==========================================</sub> | `Declare PtrSafe Function UTF8IsValid Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
UTF8
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `UTF8RuneCount` | `Declare PtrSafe Function UTF8RuneCount Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `UTF8ByteCount` | `Declare PtrSafe Function UTF8ByteCount Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

---

## filecompression

| Hàm | Khai báo VBA |
|---|---|
| `GzipCompressFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function GzipCompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |

**Mô tả:** ==========================================
GZIP FILE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcPath` | LongPtr |
| `dstPath` | LongPtr |

**Trả về:** Long

| `GzipDecompressFile` | `Declare PtrSafe Function GzipDecompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcPath` | LongPtr |
| `dstPath` | LongPtr |

**Trả về:** Long

| `ZlibCompressFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function ZlibCompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |

**Mô tả:** ==========================================
ZLIB FILE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcPath` | LongPtr |
| `dstPath` | LongPtr |

**Trả về:** Long

| `ZlibDecompressFile` | `Declare PtrSafe Function ZlibDecompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcPath` | LongPtr |
| `dstPath` | LongPtr |

**Trả về:** Long

---

## fileio

| Hàm | Khai báo VBA |
|---|---|
| `GT_FreeString`<br><sub>GT_FreeString: Giải phóng vùng nhớ mà Go đã cấp phát khi ...</sub> | `Declare PtrSafe Sub GT_FreeString Lib "GoToolkit.dll" (ByVal s As LongPtr)` |

**Mô tả:** GT_FreeString: Giải phóng vùng nhớ mà Go đã cấp phát khi trả về chuỗi.
Phía Delphi/C++ PHẢI gọi hàm này sau khi đã sao chép xong nội dung chuỗi.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Không có (Sub)

| `GT_GetFileSize`<br><sub>GT_GetFileSize: Lấy kích thước file (bytes). Thất bại trả...</sub> | `Declare PtrSafe Function GT_GetFileSize Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongLong` |

**Mô tả:** GT_GetFileSize: Lấy kích thước file (bytes). Thất bại trả về -1.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** LongLong

| `GT_FileExists`<br><sub>GT_FileExists: Kiểm tra file có tồn tại không. (1: Có, 0:...</sub> | `Declare PtrSafe Function GT_FileExists Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** GT_FileExists: Kiểm tra file có tồn tại không. (1: Có, 0: Không)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `GT_FolderExists`<br><sub>GT_FolderExists: Kiểm tra thư mục có tồn tại không. (1: C...</sub> | `Declare PtrSafe Function GT_FolderExists Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** GT_FolderExists: Kiểm tra thư mục có tồn tại không. (1: Có, 0: Không)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `GT_GetPathType`<br><sub>GT_GetPathType: Phân biệt đường dẫn. (1: File, 2: Folder,...</sub> | `Declare PtrSafe Function GT_GetPathType Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** GT_GetPathType: Phân biệt đường dẫn. (1: File, 2: Folder, 0: Không tồn tại)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `GT_ReadFile`<br><sub>GT_ReadFile: Đọc toàn bộ nội dung file trả về chuỗi.</sub> | `Declare PtrSafe Function GT_ReadFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |

**Mô tả:** GT_ReadFile: Đọc toàn bộ nội dung file trả về chuỗi.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** LongPtr

| `GT_WriteFile`<br><sub>GT_WriteFile: Ghi đè nội dung vào file. (1: Thành công, 0...</sub> | `Declare PtrSafe Function GT_WriteFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal content As LongPtr) As Long` |

**Mô tả:** GT_WriteFile: Ghi đè nội dung vào file. (1: Thành công, 0: Thất bại)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |
| `content` | LongPtr |

**Trả về:** Long

| `GT_AppendFile`<br><sub>GT_AppendFile: Ghi thêm vào cuối file. (1: Thành công, 0:...</sub> | `Declare PtrSafe Function GT_AppendFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal content As LongPtr) As Long` |

**Mô tả:** GT_AppendFile: Ghi thêm vào cuối file. (1: Thành công, 0: Thất bại)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |
| `content` | LongPtr |

**Trả về:** Long

| `GT_CopyFile`<br><sub>GT_CopyFile: Sao chép file. (1: Thành công, 0: Thất bại)</sub> | `Declare PtrSafe Function GT_CopyFile Lib "GoToolkit.dll" (ByVal sSrc As LongPtr, ByVal sDst As LongPtr) As Long` |

**Mô tả:** GT_CopyFile: Sao chép file. (1: Thành công, 0: Thất bại)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `sSrc` | LongPtr |
| `sDst` | LongPtr |

**Trả về:** Long

| `GT_MoveFile`<br><sub>GT_MoveFile: Di chuyển file hoặc đổi tên.</sub> | `Declare PtrSafe Function GT_MoveFile Lib "GoToolkit.dll" (ByVal sSrc As LongPtr, ByVal sDst As LongPtr) As Long` |

**Mô tả:** GT_MoveFile: Di chuyển file hoặc đổi tên.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `sSrc` | LongPtr |
| `sDst` | LongPtr |

**Trả về:** Long

| `GT_DeleteFile`<br><sub>GT_DeleteFile: Xóa file.</sub> | `Declare PtrSafe Function GT_DeleteFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** GT_DeleteFile: Xóa file.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `GT_CreateFolder`<br><sub>GT_CreateFolder: Tạo thư mục (bao gồm cả thư mục cha nếu ...</sub> | `Declare PtrSafe Function GT_CreateFolder Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** GT_CreateFolder: Tạo thư mục (bao gồm cả thư mục cha nếu chưa có).

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `GT_DeleteFolder`<br><sub>GT_DeleteFolder: Xóa thư mục và tất cả nội dung bên trong.</sub> | `Declare PtrSafe Function GT_DeleteFolder Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** GT_DeleteFolder: Xóa thư mục và tất cả nội dung bên trong.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `GT_ClearFolder`<br><sub>GT_ClearFolder: Xóa sạch nội dung trong thư mục nhưng giữ...</sub> | `Declare PtrSafe Function GT_ClearFolder Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** GT_ClearFolder: Xóa sạch nội dung trong thư mục nhưng giữ lại thư mục đó.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `GT_ListFiles`<br><sub>GT_ListFiles: Liệt kê file trong thư mục, cách nhau bằng ...</sub> | `Declare PtrSafe Function GT_ListFiles Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |

**Mô tả:** GT_ListFiles: Liệt kê file trong thư mục, cách nhau bằng dấu |

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** LongPtr

| `GT_ListFolders`<br><sub>GT_ListFolders: Liệt kê thư mục con, cách nhau bằng dấu |</sub> | `Declare PtrSafe Function GT_ListFolders Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |

**Mô tả:** GT_ListFolders: Liệt kê thư mục con, cách nhau bằng dấu |

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** LongPtr

| `GT_GetCurrentDir`<br><sub>GT_GetCurrentDir: Lấy thư mục đang chạy ứng dụng.</sub> | `Declare PtrSafe Function GT_GetCurrentDir Lib "GoToolkit.dll" As LongPtr` |

**Mô tả:** GT_GetCurrentDir: Lấy thư mục đang chạy ứng dụng.

**Trả về:** LongPtr

| `GT_GetTempDir`<br><sub>GT_GetTempDir: Lấy đường dẫn thư mục Temp của Windows.</sub> | `Declare PtrSafe Function GT_GetTempDir Lib "GoToolkit.dll" As LongPtr` |

**Mô tả:** GT_GetTempDir: Lấy đường dẫn thư mục Temp của Windows.

**Trả về:** LongPtr

| `GT_GetFileModTime`<br><sub>GT_GetFileModTime: Lấy ngày giờ cập nhật cuối (Định dạng:...</sub> | `Declare PtrSafe Function GT_GetFileModTime Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |

**Mô tả:** GT_GetFileModTime: Lấy ngày giờ cập nhật cuối (Định dạng: YYYY-MM-DD HH:MM:SS)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** LongPtr

| `GT_GetFileExtension`<br><sub>GT_GetFileExtension: Lấy phần mở rộng (ví dụ: .exe, .txt)</sub> | `Declare PtrSafe Function GT_GetFileExtension Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |

**Mô tả:** GT_GetFileExtension: Lấy phần mở rộng (ví dụ: .exe, .txt)

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** LongPtr

| `GT_GetFileName`<br><sub>GT_GetFileName: Lấy tên file kèm đuôi từ đường dẫn đầy đủ.</sub> | `Declare PtrSafe Function GT_GetFileName Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |

**Mô tả:** GT_GetFileName: Lấy tên file kèm đuôi từ đường dẫn đầy đủ.

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** LongPtr

---

## hash

| Hàm | Khai báo VBA |
|---|---|
| `MD5Hash` | `Declare PtrSafe Function MD5Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `input` | LongPtr |

**Trả về:** LongPtr

| `SHA1Hash` | `Declare PtrSafe Function SHA1Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `input` | LongPtr |

**Trả về:** LongPtr

| `SHA256Hash` | `Declare PtrSafe Function SHA256Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `input` | LongPtr |

**Trả về:** LongPtr

| `SHA512Hash` | `Declare PtrSafe Function SHA512Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `input` | LongPtr |

**Trả về:** LongPtr

---

## http

| Hàm | Khai báo VBA |
|---|---|
| `HttpSetTimeout`<br><sub>==========================================</sub> | `Declare PtrSafe Sub HttpSetTimeout Lib "GoToolkit.dll" (ByVal seconds As Long)` |

**Mô tả:** ==========================================
SET TIMEOUT
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `seconds` | Long |

**Trả về:** Không có (Sub)

| `HttpGet`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpGet Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
GET
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |

**Trả về:** LongPtr

| `HttpGetWithHeader` | `Declare PtrSafe Function HttpGetWithHeader Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal headerKey As LongPtr, ByVal headerVal As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |
| `headerKey` | LongPtr |
| `headerVal` | LongPtr |

**Trả về:** LongPtr

| `HttpPostJSON`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpPostJSON Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal jsonBody As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
POST
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |
| `jsonBody` | LongPtr |

**Trả về:** LongPtr

| `HttpPostForm` | `Declare PtrSafe Function HttpPostForm Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal formData As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |
| `formData` | LongPtr |

**Trả về:** LongPtr

| `HttpPut`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpPut Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal jsonBody As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
PUT / PATCH / DELETE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |
| `jsonBody` | LongPtr |

**Trả về:** LongPtr

| `HttpPatch` | `Declare PtrSafe Function HttpPatch Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal jsonBody As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |
| `jsonBody` | LongPtr |

**Trả về:** LongPtr

| `HttpDelete` | `Declare PtrSafe Function HttpDelete Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |

**Trả về:** LongPtr

| `HttpGetStatusCode`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpGetStatusCode Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr) As Long` |

**Mô tả:** ==========================================
STATUS CODE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |

**Trả về:** Long

| `HttpDownloadFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpDownloadFile Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal dstPath As LongPtr) As Long` |

**Mô tả:** ==========================================
DOWNLOAD FILE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `reqUrl` | LongPtr |
| `dstPath` | LongPtr |

**Trả về:** Long

| `JsonGetValue`<br><sub>==========================================</sub> | `Declare PtrSafe Function JsonGetValue Lib "GoToolkit.dll" (ByVal jsonStr As LongPtr, ByVal key As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
JSON HELPER
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `jsonStr` | LongPtr |
| `key` | LongPtr |

**Trả về:** LongPtr

| `JsonBuildObject` | `Declare PtrSafe Function JsonBuildObject Lib "GoToolkit.dll" (ByVal kvPairs As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `kvPairs` | LongPtr |

**Trả về:** LongPtr

---

## jwt

| Hàm | Khai báo VBA |
|---|---|
| `JwtGenerateKeys`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtGenerateKeys Lib "GoToolkit.dll" (ByVal privatePath As LongPtr, ByVal publicPath As LongPtr) As Long` |

**Mô tả:** ==========================================
TẠO / NẠP RSA KEY
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `privatePath` | LongPtr |
| `publicPath` | LongPtr |

**Trả về:** Long

| `JwtLoadPrivateKey` | `Declare PtrSafe Function JwtLoadPrivateKey Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `JwtLoadPublicKey` | `Declare PtrSafe Function JwtLoadPublicKey Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `fPath` | LongPtr |

**Trả về:** Long

| `JwtCreateToken`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtCreateToken Lib "GoToolkit.dll" (ByVal sub As LongPtr, ByVal role As LongPtr, ByVal iss As LongPtr, ByVal aud As LongPtr, ByVal expSeconds As Long) As LongPtr` |

**Mô tả:** ==========================================
TẠO TOKEN
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `sub` | LongPtr |
| `role` | LongPtr |
| `iss` | LongPtr |
| `aud` | LongPtr |
| `expSeconds` | Long |

**Trả về:** LongPtr

| `JwtVerifyToken`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtVerifyToken Lib "GoToolkit.dll" (ByVal token As LongPtr) As Long` |

**Mô tả:** ==========================================
XÁC THỰC TOKEN
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `token` | LongPtr |

**Trả về:** Long

| `JwtGetClaims`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtGetClaims Lib "GoToolkit.dll" (ByVal token As LongPtr) As LongPtr` |

**Mô tả:** ==========================================
LẤY THÔNG TIN TỪ TOKEN
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `token` | LongPtr |

**Trả về:** LongPtr

| `JwtGetField` | `Declare PtrSafe Function JwtGetField Lib "GoToolkit.dll" (ByVal token As LongPtr, ByVal field As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `token` | LongPtr |
| `field` | LongPtr |

**Trả về:** LongPtr

| `JwtIsExpired` | `Declare PtrSafe Function JwtIsExpired Lib "GoToolkit.dll" (ByVal token As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `token` | LongPtr |

**Trả về:** Long

| `JwtGetExpireTime` | `Declare PtrSafe Function JwtGetExpireTime Lib "GoToolkit.dll" (ByVal token As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `token` | LongPtr |

**Trả về:** LongPtr

| `JwtRefreshToken`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtRefreshToken Lib "GoToolkit.dll" (ByVal token As LongPtr, ByVal expSeconds As Long) As LongPtr` |

**Mô tả:** ==========================================
LÀM MỚI TOKEN
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `token` | LongPtr |
| `expSeconds` | Long |

**Trả về:** LongPtr

---

## stringutils

| Hàm | Khai báo VBA |
|---|---|
| `StrToUpper` | `Declare PtrSafe Function StrToUpper Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `StrToLower` | `Declare PtrSafe Function StrToLower Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `StrTrim` | `Declare PtrSafe Function StrTrim Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `StrReplace` | `Declare PtrSafe Function StrReplace Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal old As LongPtr, ByVal new As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `old` | LongPtr |
| `new` | LongPtr |

**Trả về:** LongPtr

| `StrContains` | `Declare PtrSafe Function StrContains Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal sub As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `sub` | LongPtr |

**Trả về:** Long

| `StrStartsWith` | `Declare PtrSafe Function StrStartsWith Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal prefix As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `prefix` | LongPtr |

**Trả về:** Long

| `StrEndsWith` | `Declare PtrSafe Function StrEndsWith Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal suffix As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `suffix` | LongPtr |

**Trả về:** Long

| `StrLength` | `Declare PtrSafe Function StrLength Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `StrReverse` | `Declare PtrSafe Function StrReverse Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** LongPtr

| `StrCount` | `Declare PtrSafe Function StrCount Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal sub As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `sub` | LongPtr |

**Trả về:** Long

| `StrIsNumeric` | `Declare PtrSafe Function StrIsNumeric Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `StrIsPalindrome` | `Declare PtrSafe Function StrIsPalindrome Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

---

## validation

| Hàm | Khai báo VBA |
|---|---|
| `ValidateEmail`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateEmail Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
EMAIL
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidatePhone`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidatePhone Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
PHONE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidatePhoneVN` | `Declare PtrSafe Function ValidatePhoneVN Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateURL`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateURL Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
URL
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateIPv4`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateIPv4 Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
IP
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateIPv6` | `Declare PtrSafe Function ValidateIPv6 Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateIP` | `Declare PtrSafe Function ValidateIP Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateInteger`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateInteger Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
SỐ
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateFloat` | `Declare PtrSafe Function ValidateFloat Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateInRange` | `Declare PtrSafe Function ValidateInRange Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal nMin As Double, ByVal nMax As Double) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `nMin` | Double |
| `nMax` | Double |

**Trả về:** Long

| `ValidateMinLength`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateMinLength Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal nMin As Long) As Long` |

**Mô tả:** ==========================================
CHUỖI
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `nMin` | Long |

**Trả về:** Long

| `ValidateMaxLength` | `Declare PtrSafe Function ValidateMaxLength Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal nMax As Long) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `nMax` | Long |

**Trả về:** Long

| `ValidateNotEmpty` | `Declare PtrSafe Function ValidateNotEmpty Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateAlpha` | `Declare PtrSafe Function ValidateAlpha Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateAlphaNumeric` | `Declare PtrSafe Function ValidateAlphaNumeric Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateContainsUpper` | `Declare PtrSafe Function ValidateContainsUpper Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateContainsLower` | `Declare PtrSafe Function ValidateContainsLower Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateContainsDigit` | `Declare PtrSafe Function ValidateContainsDigit Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateContainsSpecial` | `Declare PtrSafe Function ValidateContainsSpecial Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidatePassword`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidatePassword Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
PASSWORD
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidatePasswordStrength` | `Declare PtrSafe Function ValidatePasswordStrength Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateDate`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateDate Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal fmt As LongPtr) As Long` |

**Mô tả:** ==========================================
DATE
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `fmt` | LongPtr |

**Trả về:** Long

| `ValidateCreditCard`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateCreditCard Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

**Mô tả:** ==========================================
CREDIT CARD (Luhn)
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |

**Trả về:** Long

| `ValidateRegex`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateRegex Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal pattern As LongPtr) As Long` |

**Mô tả:** ==========================================
REGEX CUSTOM
==========================================

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `s` | LongPtr |
| `pattern` | LongPtr |

**Trả về:** Long

---

## zipfile

| Hàm | Khai báo VBA |
|---|---|
| `ZipCompressFile` | `Declare PtrSafe Function ZipCompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstZip As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcPath` | LongPtr |
| `dstZip` | LongPtr |

**Trả về:** Long

| `ZipCompressFolder` | `Declare PtrSafe Function ZipCompressFolder Lib "GoToolkit.dll" (ByVal srcFolder As LongPtr, ByVal dstZip As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcFolder` | LongPtr |
| `dstZip` | LongPtr |

**Trả về:** Long

| `ZipDecompress` | `Declare PtrSafe Function ZipDecompress Lib "GoToolkit.dll" (ByVal srcZip As LongPtr, ByVal dstFolder As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcZip` | LongPtr |
| `dstFolder` | LongPtr |

**Trả về:** Long

| `ZipAddFile` | `Declare PtrSafe Function ZipAddFile Lib "GoToolkit.dll" (ByVal existingZip As LongPtr, ByVal newFile As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `existingZip` | LongPtr |
| `newFile` | LongPtr |

**Trả về:** Long

| `ZipListFiles` | `Declare PtrSafe Function ZipListFiles Lib "GoToolkit.dll" (ByVal srcZip As LongPtr) As LongPtr` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcZip` | LongPtr |

**Trả về:** LongPtr

| `ZipGetFileCount` | `Declare PtrSafe Function ZipGetFileCount Lib "GoToolkit.dll" (ByVal srcZip As LongPtr) As Long` |

**Mô tả:** _Chưa có mô tả._

**Tham số:**

| Tên | Kiểu VBA |
|---|---|
| `srcZip` | LongPtr |

**Trả về:** Long

---

_Sinh bởi **GoDocGen** `Mode=vba` — 05/05/2026 16:14_
