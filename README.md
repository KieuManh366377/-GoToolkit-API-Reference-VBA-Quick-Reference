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
| `BcryptVerifyPassword` | `Declare PtrSafe Function BcryptVerifyPassword Lib "GoToolkit.dll" (ByVal password As LongPtr, ByVal hashed As LongPtr) As Long` |
| `BcryptEncryptString`<br><sub>==========================================</sub> | `Declare PtrSafe Function BcryptEncryptString Lib "GoToolkit.dll" (ByVal plaintext As LongPtr, ByVal password As LongPtr) As LongPtr` |
| `BcryptDecryptString` | `Declare PtrSafe Function BcryptDecryptString Lib "GoToolkit.dll" (ByVal encrypted As LongPtr, ByVal password As LongPtr) As LongPtr` |
| `BcryptEncryptFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function BcryptEncryptFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr, ByVal password As LongPtr) As Long` |
| `BcryptDecryptFile` | `Declare PtrSafe Function BcryptDecryptFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr, ByVal password As LongPtr) As Long` |
| `BcryptSignString`<br><sub>==========================================</sub> | `Declare PtrSafe Function BcryptSignString Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal secretKey As LongPtr) As LongPtr` |
| `BcryptVerifyString` | `Declare PtrSafe Function BcryptVerifyString Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal signature As LongPtr, ByVal secretKey As LongPtr) As Long` |
| `BcryptSignFile` | `Declare PtrSafe Function BcryptSignFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal secretKey As LongPtr) As LongPtr` |
| `BcryptVerifyFile` | `Declare PtrSafe Function BcryptVerifyFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal signature As LongPtr, ByVal secretKey As LongPtr) As Long` |
| `BcryptFreeString` | `Declare PtrSafe Sub BcryptFreeString Lib "GoToolkit.dll" (ByVal s As LongPtr)` |

---

## compression

| Hàm | Khai báo VBA |
|---|---|
| `GzipCompress`<br><sub>==========================================</sub> | `Declare PtrSafe Function GzipCompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `GzipDecompress` | `Declare PtrSafe Function GzipDecompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `ZlibCompress`<br><sub>==========================================</sub> | `Declare PtrSafe Function ZlibCompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `ZlibDecompress` | `Declare PtrSafe Function ZlibDecompress Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `CompressRatio`<br><sub>==========================================</sub> | `Declare PtrSafe Function CompressRatio Lib "GoToolkit.dll" (ByVal original As LongPtr, ByVal compressed As LongPtr) As Long` |
| `FreeString` | `Declare PtrSafe Sub FreeString Lib "GoToolkit.dll" (ByVal s As LongPtr)` |

---

## converter

| Hàm | Khai báo VBA |
|---|---|
| `GT_Mem_Copy`<br><sub>GT_Mem_Copy: Thay the CopyMemory cua kernel32</sub> | `Declare PtrSafe Sub GT_Mem_Copy Lib "GoToolkit.dll" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)` |
| `GT_Mem_PtrToStringUTF8`<br><sub>GT_Mem_PtrToStringUTF8: Sao chep chuoi null-terminated tu...</sub> | `Declare PtrSafe Function GT_Mem_PtrToStringUTF8 Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |
| `GT_Mem_Free`<br><sub>GT_Mem_Free: Giai phong bo nho tu GT_Mem_PtrToStringUTF8 ...</sub> | `Declare PtrSafe Sub GT_Mem_Free Lib "GoToolkit.dll" (ByVal ptr As LongPtr)` |
| `GT_Mem_PtrReadByte`<br><sub>GT_Mem_PtrReadByte: Doc 1 byte tai dia chi chi dinh</sub> | `Declare PtrSafe Function GT_Mem_PtrReadByte Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |
| `GT_Mem_PtrReadInt32`<br><sub>GT_Mem_PtrReadInt32: Doc 4 byte (Long trong VBA)</sub> | `Declare PtrSafe Function GT_Mem_PtrReadInt32 Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As Long` |
| `GT_Mem_PtrReadInt64`<br><sub>GT_Mem_PtrReadInt64: Doc 8 byte (LongLong trong VBA 64-bit)</sub> | `Declare PtrSafe Function GT_Mem_PtrReadInt64 Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |
| `GT_Mem_PtrReadDouble`<br><sub>GT_Mem_PtrReadDouble: Doc 8 byte (Double trong VBA)</sub> | `Declare PtrSafe Function GT_Mem_PtrReadDouble Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As Double` |
| `GT_Mem_Zero`<br><sub>GT_Mem_Zero: Xoa trang vung nho (huu ich cho du lieu nhay...</sub> | `Declare PtrSafe Sub GT_Mem_Zero Lib "GoToolkit.dll" (ByVal ptr As LongPtr, ByVal length As LongPtr)` |
| `GT_PtrToString`<br><sub>GT_PtrToString: Ten ngan hon cua GT_Mem_PtrToStringUTF8</sub> | `Declare PtrSafe Function GT_PtrToString Lib "GoToolkit.dll" (ByVal ptr As LongPtr) As LongPtr` |

---

## crypto

| Hàm | Khai báo VBA |
|---|---|
| `SetKey` | `Declare PtrSafe Function SetKey Lib "GoToolkit.dll" (ByVal key As LongPtr) As Long` |
| `Encrypt` | `Declare PtrSafe Function Encrypt Lib "GoToolkit.dll" (ByVal plaintext As LongPtr) As LongPtr` |
| `Decrypt` | `Declare PtrSafe Function Decrypt Lib "GoToolkit.dll" (ByVal ciphertext As LongPtr) As LongPtr` |

---

## datetime

| Hàm | Khai báo VBA |
|---|---|
| `DateTimeNow`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeNow Lib "GoToolkit.dll" As LongPtr` |
| `DateNow` | `Declare PtrSafe Function DateNow Lib "GoToolkit.dll" As LongPtr` |
| `TimeNow` | `Declare PtrSafe Function TimeNow Lib "GoToolkit.dll" As LongPtr` |
| `DateTimeUTC` | `Declare PtrSafe Function DateTimeUTC Lib "GoToolkit.dll" As LongPtr` |
| `UnixTimestamp` | `Declare PtrSafe Function UnixTimestamp Lib "GoToolkit.dll" As LongLong` |
| `UnixTimestampMs` | `Declare PtrSafe Function UnixTimestampMs Lib "GoToolkit.dll" As LongLong` |
| `DateTimeFormat`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeFormat Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal fmtIn As LongPtr, ByVal fmtOut As LongPtr) As LongPtr` |
| `UnixToDateTime` | `Declare PtrSafe Function UnixToDateTime Lib "GoToolkit.dll" (ByVal unix As LongLong) As LongPtr` |
| `DateTimeToUnix` | `Declare PtrSafe Function DateTimeToUnix Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As LongLong` |
| `DateTimeAddDays`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeAddDays Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nDays As Long) As LongPtr` |
| `DateTimeAddMonths` | `Declare PtrSafe Function DateTimeAddMonths Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nMonths As Long) As LongPtr` |
| `DateTimeAddYears` | `Declare PtrSafe Function DateTimeAddYears Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nYears As Long) As LongPtr` |
| `DateTimeAddHours` | `Declare PtrSafe Function DateTimeAddHours Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nHours As Long) As LongPtr` |
| `DateTimeAddMinutes` | `Declare PtrSafe Function DateTimeAddMinutes Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal nMins As Long) As LongPtr` |
| `DateTimeDiffDays`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeDiffDays Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |
| `DateTimeDiffHours` | `Declare PtrSafe Function DateTimeDiffHours Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |
| `DateTimeDiffMinutes` | `Declare PtrSafe Function DateTimeDiffMinutes Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |
| `DateTimeDiffSeconds` | `Declare PtrSafe Function DateTimeDiffSeconds Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |
| `DateTimeIsBefore` | `Declare PtrSafe Function DateTimeIsBefore Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |
| `DateTimeIsAfter` | `Declare PtrSafe Function DateTimeIsAfter Lib "GoToolkit.dll" (ByVal dtStr1 As LongPtr, ByVal dtStr2 As LongPtr) As Long` |
| `DateGetYear`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateGetYear Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |
| `DateGetMonth` | `Declare PtrSafe Function DateGetMonth Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |
| `DateGetDay` | `Declare PtrSafe Function DateGetDay Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |
| `DateGetWeekday` | `Declare PtrSafe Function DateGetWeekday Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As LongPtr` |
| `DateGetWeekNumber` | `Declare PtrSafe Function DateGetWeekNumber Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |
| `DateGetDayOfYear` | `Declare PtrSafe Function DateGetDayOfYear Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |
| `DateIsLeapYear` | `Declare PtrSafe Function DateIsLeapYear Lib "GoToolkit.dll" (ByVal nYear As Long) As Long` |
| `DateGetDaysInMonth` | `Declare PtrSafe Function DateGetDaysInMonth Lib "GoToolkit.dll" (ByVal nYear As Long, ByVal nMonth As Long) As Long` |
| `DateGetQuarter` | `Declare PtrSafe Function DateGetQuarter Lib "GoToolkit.dll" (ByVal dtStr As LongPtr) As Long` |
| `DateTimeConvertTZ`<br><sub>==========================================</sub> | `Declare PtrSafe Function DateTimeConvertTZ Lib "GoToolkit.dll" (ByVal dtStr As LongPtr, ByVal fromTZ As LongPtr, ByVal toTZ As LongPtr) As LongPtr` |

---

## encodeutils

| Hàm | Khai báo VBA |
|---|---|
| `Base64Encode`<br><sub>==========================================</sub> | `Declare PtrSafe Function Base64Encode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `Base64Decode` | `Declare PtrSafe Function Base64Decode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `Base64EncodeURL` | `Declare PtrSafe Function Base64EncodeURL Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `Base64DecodeURL` | `Declare PtrSafe Function Base64DecodeURL Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `Base64IsValid` | `Declare PtrSafe Function Base64IsValid Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `HexEncode`<br><sub>==========================================</sub> | `Declare PtrSafe Function HexEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `HexDecode` | `Declare PtrSafe Function HexDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `HexIsValid` | `Declare PtrSafe Function HexIsValid Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `HexToDecimal` | `Declare PtrSafe Function HexToDecimal Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongLong` |
| `DecimalToHex` | `Declare PtrSafe Function DecimalToHex Lib "GoToolkit.dll" (ByVal n As LongLong) As LongPtr` |
| `DecimalToBinary` | `Declare PtrSafe Function DecimalToBinary Lib "GoToolkit.dll" (ByVal n As LongLong) As LongPtr` |
| `BinaryToDecimal` | `Declare PtrSafe Function BinaryToDecimal Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongLong` |
| `DecimalToOctal` | `Declare PtrSafe Function DecimalToOctal Lib "GoToolkit.dll" (ByVal n As LongLong) As LongPtr` |
| `OctalToDecimal` | `Declare PtrSafe Function OctalToDecimal Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongLong` |
| `URLEncode`<br><sub>==========================================</sub> | `Declare PtrSafe Function URLEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `URLDecode` | `Declare PtrSafe Function URLDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `URLPathEncode` | `Declare PtrSafe Function URLPathEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `URLPathDecode` | `Declare PtrSafe Function URLPathDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `HTMLEncode`<br><sub>==========================================</sub> | `Declare PtrSafe Function HTMLEncode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `HTMLDecode` | `Declare PtrSafe Function HTMLDecode Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `UTF8IsValid`<br><sub>==========================================</sub> | `Declare PtrSafe Function UTF8IsValid Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `UTF8RuneCount` | `Declare PtrSafe Function UTF8RuneCount Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `UTF8ByteCount` | `Declare PtrSafe Function UTF8ByteCount Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

---

## filecompression

| Hàm | Khai báo VBA |
|---|---|
| `GzipCompressFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function GzipCompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |
| `GzipDecompressFile` | `Declare PtrSafe Function GzipDecompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |
| `ZlibCompressFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function ZlibCompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |
| `ZlibDecompressFile` | `Declare PtrSafe Function ZlibDecompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstPath As LongPtr) As Long` |

---

## fileio

| Hàm | Khai báo VBA |
|---|---|
| `GT_FreeString`<br><sub>GT_FreeString: Giải phóng vùng nhớ mà Go đã cấp phát khi ...</sub> | `Declare PtrSafe Sub GT_FreeString Lib "GoToolkit.dll" (ByVal s As LongPtr)` |
| `GT_GetFileSize`<br><sub>GT_GetFileSize: Lấy kích thước file (bytes). Thất bại trả...</sub> | `Declare PtrSafe Function GT_GetFileSize Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongLong` |
| `GT_FileExists`<br><sub>GT_FileExists: Kiểm tra file có tồn tại không. (1: Có, 0:...</sub> | `Declare PtrSafe Function GT_FileExists Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `GT_FolderExists`<br><sub>GT_FolderExists: Kiểm tra thư mục có tồn tại không. (1: C...</sub> | `Declare PtrSafe Function GT_FolderExists Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `GT_GetPathType`<br><sub>GT_GetPathType: Phân biệt đường dẫn. (1: File, 2: Folder,...</sub> | `Declare PtrSafe Function GT_GetPathType Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `GT_ReadFile`<br><sub>GT_ReadFile: Đọc toàn bộ nội dung file trả về chuỗi.</sub> | `Declare PtrSafe Function GT_ReadFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |
| `GT_WriteFile`<br><sub>GT_WriteFile: Ghi đè nội dung vào file. (1: Thành công, 0...</sub> | `Declare PtrSafe Function GT_WriteFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal content As LongPtr) As Long` |
| `GT_AppendFile`<br><sub>GT_AppendFile: Ghi thêm vào cuối file. (1: Thành công, 0:...</sub> | `Declare PtrSafe Function GT_AppendFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr, ByVal content As LongPtr) As Long` |
| `GT_CopyFile`<br><sub>GT_CopyFile: Sao chép file. (1: Thành công, 0: Thất bại)</sub> | `Declare PtrSafe Function GT_CopyFile Lib "GoToolkit.dll" (ByVal sSrc As LongPtr, ByVal sDst As LongPtr) As Long` |
| `GT_MoveFile`<br><sub>GT_MoveFile: Di chuyển file hoặc đổi tên.</sub> | `Declare PtrSafe Function GT_MoveFile Lib "GoToolkit.dll" (ByVal sSrc As LongPtr, ByVal sDst As LongPtr) As Long` |
| `GT_DeleteFile`<br><sub>GT_DeleteFile: Xóa file.</sub> | `Declare PtrSafe Function GT_DeleteFile Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `GT_CreateFolder`<br><sub>GT_CreateFolder: Tạo thư mục (bao gồm cả thư mục cha nếu ...</sub> | `Declare PtrSafe Function GT_CreateFolder Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `GT_DeleteFolder`<br><sub>GT_DeleteFolder: Xóa thư mục và tất cả nội dung bên trong.</sub> | `Declare PtrSafe Function GT_DeleteFolder Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `GT_ClearFolder`<br><sub>GT_ClearFolder: Xóa sạch nội dung trong thư mục nhưng giữ...</sub> | `Declare PtrSafe Function GT_ClearFolder Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `GT_ListFiles`<br><sub>GT_ListFiles: Liệt kê file trong thư mục, cách nhau bằng ...</sub> | `Declare PtrSafe Function GT_ListFiles Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |
| `GT_ListFolders`<br><sub>GT_ListFolders: Liệt kê thư mục con, cách nhau bằng dấu |</sub> | `Declare PtrSafe Function GT_ListFolders Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |
| `GT_GetCurrentDir`<br><sub>GT_GetCurrentDir: Lấy thư mục đang chạy ứng dụng.</sub> | `Declare PtrSafe Function GT_GetCurrentDir Lib "GoToolkit.dll" As LongPtr` |
| `GT_GetTempDir`<br><sub>GT_GetTempDir: Lấy đường dẫn thư mục Temp của Windows.</sub> | `Declare PtrSafe Function GT_GetTempDir Lib "GoToolkit.dll" As LongPtr` |
| `GT_GetFileModTime`<br><sub>GT_GetFileModTime: Lấy ngày giờ cập nhật cuối (Định dạng:...</sub> | `Declare PtrSafe Function GT_GetFileModTime Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |
| `GT_GetFileExtension`<br><sub>GT_GetFileExtension: Lấy phần mở rộng (ví dụ: .exe, .txt)</sub> | `Declare PtrSafe Function GT_GetFileExtension Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |
| `GT_GetFileName`<br><sub>GT_GetFileName: Lấy tên file kèm đuôi từ đường dẫn đầy đủ.</sub> | `Declare PtrSafe Function GT_GetFileName Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As LongPtr` |

---

## hash

| Hàm | Khai báo VBA |
|---|---|
| `MD5Hash` | `Declare PtrSafe Function MD5Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |
| `SHA1Hash` | `Declare PtrSafe Function SHA1Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |
| `SHA256Hash` | `Declare PtrSafe Function SHA256Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |
| `SHA512Hash` | `Declare PtrSafe Function SHA512Hash Lib "GoToolkit.dll" (ByVal input As LongPtr) As LongPtr` |

---

## http

| Hàm | Khai báo VBA |
|---|---|
| `HttpSetTimeout`<br><sub>==========================================</sub> | `Declare PtrSafe Sub HttpSetTimeout Lib "GoToolkit.dll" (ByVal seconds As Long)` |
| `HttpGet`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpGet Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr) As LongPtr` |
| `HttpGetWithHeader` | `Declare PtrSafe Function HttpGetWithHeader Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal headerKey As LongPtr, ByVal headerVal As LongPtr) As LongPtr` |
| `HttpPostJSON`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpPostJSON Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal jsonBody As LongPtr) As LongPtr` |
| `HttpPostForm` | `Declare PtrSafe Function HttpPostForm Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal formData As LongPtr) As LongPtr` |
| `HttpPut`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpPut Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal jsonBody As LongPtr) As LongPtr` |
| `HttpPatch` | `Declare PtrSafe Function HttpPatch Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal jsonBody As LongPtr) As LongPtr` |
| `HttpDelete` | `Declare PtrSafe Function HttpDelete Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr) As LongPtr` |
| `HttpGetStatusCode`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpGetStatusCode Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr) As Long` |
| `HttpDownloadFile`<br><sub>==========================================</sub> | `Declare PtrSafe Function HttpDownloadFile Lib "GoToolkit.dll" (ByVal reqUrl As LongPtr, ByVal dstPath As LongPtr) As Long` |
| `JsonGetValue`<br><sub>==========================================</sub> | `Declare PtrSafe Function JsonGetValue Lib "GoToolkit.dll" (ByVal jsonStr As LongPtr, ByVal key As LongPtr) As LongPtr` |
| `JsonBuildObject` | `Declare PtrSafe Function JsonBuildObject Lib "GoToolkit.dll" (ByVal kvPairs As LongPtr) As LongPtr` |

---

## jwt

| Hàm | Khai báo VBA |
|---|---|
| `JwtGenerateKeys`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtGenerateKeys Lib "GoToolkit.dll" (ByVal privatePath As LongPtr, ByVal publicPath As LongPtr) As Long` |
| `JwtLoadPrivateKey` | `Declare PtrSafe Function JwtLoadPrivateKey Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `JwtLoadPublicKey` | `Declare PtrSafe Function JwtLoadPublicKey Lib "GoToolkit.dll" (ByVal fPath As LongPtr) As Long` |
| `JwtCreateToken`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtCreateToken Lib "GoToolkit.dll" (ByVal sub As LongPtr, ByVal role As LongPtr, ByVal iss As LongPtr, ByVal aud As LongPtr, ByVal expSeconds As Long) As LongPtr` |
| `JwtVerifyToken`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtVerifyToken Lib "GoToolkit.dll" (ByVal token As LongPtr) As Long` |
| `JwtGetClaims`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtGetClaims Lib "GoToolkit.dll" (ByVal token As LongPtr) As LongPtr` |
| `JwtGetField` | `Declare PtrSafe Function JwtGetField Lib "GoToolkit.dll" (ByVal token As LongPtr, ByVal field As LongPtr) As LongPtr` |
| `JwtIsExpired` | `Declare PtrSafe Function JwtIsExpired Lib "GoToolkit.dll" (ByVal token As LongPtr) As Long` |
| `JwtGetExpireTime` | `Declare PtrSafe Function JwtGetExpireTime Lib "GoToolkit.dll" (ByVal token As LongPtr) As LongPtr` |
| `JwtRefreshToken`<br><sub>==========================================</sub> | `Declare PtrSafe Function JwtRefreshToken Lib "GoToolkit.dll" (ByVal token As LongPtr, ByVal expSeconds As Long) As LongPtr` |

---

## stringutils

| Hàm | Khai báo VBA |
|---|---|
| `StrToUpper` | `Declare PtrSafe Function StrToUpper Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `StrToLower` | `Declare PtrSafe Function StrToLower Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `StrTrim` | `Declare PtrSafe Function StrTrim Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `StrReplace` | `Declare PtrSafe Function StrReplace Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal old As LongPtr, ByVal new As LongPtr) As LongPtr` |
| `StrContains` | `Declare PtrSafe Function StrContains Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal sub As LongPtr) As Long` |
| `StrStartsWith` | `Declare PtrSafe Function StrStartsWith Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal prefix As LongPtr) As Long` |
| `StrEndsWith` | `Declare PtrSafe Function StrEndsWith Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal suffix As LongPtr) As Long` |
| `StrLength` | `Declare PtrSafe Function StrLength Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `StrReverse` | `Declare PtrSafe Function StrReverse Lib "GoToolkit.dll" (ByVal s As LongPtr) As LongPtr` |
| `StrCount` | `Declare PtrSafe Function StrCount Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal sub As LongPtr) As Long` |
| `StrIsNumeric` | `Declare PtrSafe Function StrIsNumeric Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `StrIsPalindrome` | `Declare PtrSafe Function StrIsPalindrome Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |

---

## validation

| Hàm | Khai báo VBA |
|---|---|
| `ValidateEmail`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateEmail Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidatePhone`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidatePhone Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidatePhoneVN` | `Declare PtrSafe Function ValidatePhoneVN Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateURL`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateURL Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateIPv4`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateIPv4 Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateIPv6` | `Declare PtrSafe Function ValidateIPv6 Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateIP` | `Declare PtrSafe Function ValidateIP Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateInteger`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateInteger Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateFloat` | `Declare PtrSafe Function ValidateFloat Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateInRange` | `Declare PtrSafe Function ValidateInRange Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal nMin As Double, ByVal nMax As Double) As Long` |
| `ValidateMinLength`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateMinLength Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal nMin As Long) As Long` |
| `ValidateMaxLength` | `Declare PtrSafe Function ValidateMaxLength Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal nMax As Long) As Long` |
| `ValidateNotEmpty` | `Declare PtrSafe Function ValidateNotEmpty Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateAlpha` | `Declare PtrSafe Function ValidateAlpha Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateAlphaNumeric` | `Declare PtrSafe Function ValidateAlphaNumeric Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateContainsUpper` | `Declare PtrSafe Function ValidateContainsUpper Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateContainsLower` | `Declare PtrSafe Function ValidateContainsLower Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateContainsDigit` | `Declare PtrSafe Function ValidateContainsDigit Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateContainsSpecial` | `Declare PtrSafe Function ValidateContainsSpecial Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidatePassword`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidatePassword Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidatePasswordStrength` | `Declare PtrSafe Function ValidatePasswordStrength Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateDate`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateDate Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal fmt As LongPtr) As Long` |
| `ValidateCreditCard`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateCreditCard Lib "GoToolkit.dll" (ByVal s As LongPtr) As Long` |
| `ValidateRegex`<br><sub>==========================================</sub> | `Declare PtrSafe Function ValidateRegex Lib "GoToolkit.dll" (ByVal s As LongPtr, ByVal pattern As LongPtr) As Long` |

---

## zipfile

| Hàm | Khai báo VBA |
|---|---|
| `ZipCompressFile` | `Declare PtrSafe Function ZipCompressFile Lib "GoToolkit.dll" (ByVal srcPath As LongPtr, ByVal dstZip As LongPtr) As Long` |
| `ZipCompressFolder` | `Declare PtrSafe Function ZipCompressFolder Lib "GoToolkit.dll" (ByVal srcFolder As LongPtr, ByVal dstZip As LongPtr) As Long` |
| `ZipDecompress` | `Declare PtrSafe Function ZipDecompress Lib "GoToolkit.dll" (ByVal srcZip As LongPtr, ByVal dstFolder As LongPtr) As Long` |
| `ZipAddFile` | `Declare PtrSafe Function ZipAddFile Lib "GoToolkit.dll" (ByVal existingZip As LongPtr, ByVal newFile As LongPtr) As Long` |
| `ZipListFiles` | `Declare PtrSafe Function ZipListFiles Lib "GoToolkit.dll" (ByVal srcZip As LongPtr) As LongPtr` |
| `ZipGetFileCount` | `Declare PtrSafe Function ZipGetFileCount Lib "GoToolkit.dll" (ByVal srcZip As LongPtr) As Long` |

---

_Sinh bởi **GoDocGen** `Mode=vba` — 05/05/2026 16:02_
