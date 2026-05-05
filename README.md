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
| `BcryptHashPassword`<br><sub>==========================================</sub> | `Function BcryptHashPassword(ByVal password As String) As LongPtr` |
| `BcryptVerifyPassword` | `Function BcryptVerifyPassword(ByVal password As String, ByVal hashed As String) As Long` |
| `BcryptEncryptString`<br><sub>==========================================</sub> | `Function BcryptEncryptString(ByVal plaintext As String, ByVal password As String) As LongPtr` |
| `BcryptDecryptString` | `Function BcryptDecryptString(ByVal encrypted As String, ByVal password As String) As LongPtr` |
| `BcryptEncryptFile`<br><sub>==========================================</sub> | `Function BcryptEncryptFile(ByVal srcPath As String, ByVal dstPath As String, ByVal password As String) As Long` |
| `BcryptDecryptFile` | `Function BcryptDecryptFile(ByVal srcPath As String, ByVal dstPath As String, ByVal password As String) As Long` |
| `BcryptSignString`<br><sub>==========================================</sub> | `Function BcryptSignString(ByVal s As String, ByVal secretKey As String) As LongPtr` |
| `BcryptVerifyString` | `Function BcryptVerifyString(ByVal s As String, ByVal signature As String, ByVal secretKey As String) As Long` |
| `BcryptSignFile` | `Function BcryptSignFile(ByVal fPath As String, ByVal secretKey As String) As LongPtr` |
| `BcryptVerifyFile` | `Function BcryptVerifyFile(ByVal fPath As String, ByVal signature As String, ByVal secretKey As String) As Long` |
| `BcryptFreeString` | `Sub BcryptFreeString(ByVal s As String)` |

---

## compression

| Hàm | Khai báo VBA |
|---|---|
| `GzipCompress`<br><sub>==========================================</sub> | `Function GzipCompress(ByVal s As String) As LongPtr` |
| `GzipDecompress` | `Function GzipDecompress(ByVal s As String) As LongPtr` |
| `ZlibCompress`<br><sub>==========================================</sub> | `Function ZlibCompress(ByVal s As String) As LongPtr` |
| `ZlibDecompress` | `Function ZlibDecompress(ByVal s As String) As LongPtr` |
| `CompressRatio`<br><sub>==========================================</sub> | `Function CompressRatio(ByVal original As String, ByVal compressed As String) As Long` |
| `FreeString` | `Sub FreeString(ByVal s As String)` |

---

## converter

| Hàm | Khai báo VBA |
|---|---|
| `GT_Mem_Copy`<br><sub>GT_Mem_Copy: Thay the CopyMemory cua kernel32</sub> | `Sub GT_Mem_Copy(ByVal destination As Long, ByVal source As Long, ByVal length As Long)` |
| `GT_Mem_PtrToStringUTF8`<br><sub>GT_Mem_PtrToStringUTF8: Sao chep chuoi null-terminated tu...</sub> | `Function GT_Mem_PtrToStringUTF8(ByVal ptr As Long) As Long` |
| `GT_Mem_Free`<br><sub>GT_Mem_Free: Giai phong bo nho tu GT_Mem_PtrToStringUTF8 ...</sub> | `Sub GT_Mem_Free(ByVal ptr As Long)` |
| `GT_Mem_PtrReadByte`<br><sub>GT_Mem_PtrReadByte: Doc 1 byte tai dia chi chi dinh</sub> | `Function GT_Mem_PtrReadByte(ByVal ptr As Long) As Long` |
| `GT_Mem_PtrReadInt32`<br><sub>GT_Mem_PtrReadInt32: Doc 4 byte (Long trong VBA)</sub> | `Function GT_Mem_PtrReadInt32(ByVal ptr As Long) As Long` |
| `GT_Mem_PtrReadInt64`<br><sub>GT_Mem_PtrReadInt64: Doc 8 byte (LongLong trong VBA 64-bit)</sub> | `Function GT_Mem_PtrReadInt64(ByVal ptr As Long) As Long` |
| `GT_Mem_PtrReadDouble`<br><sub>GT_Mem_PtrReadDouble: Doc 8 byte (Double trong VBA)</sub> | `Function GT_Mem_PtrReadDouble(ByVal ptr As Long) As Double` |
| `GT_Mem_Zero`<br><sub>GT_Mem_Zero: Xoa trang vung nho (huu ich cho du lieu nhay...</sub> | `Sub GT_Mem_Zero(ByVal ptr As Long, ByVal length As Long)` |
| `GT_PtrToString`<br><sub>GT_PtrToString: Ten ngan hon cua GT_Mem_PtrToStringUTF8</sub> | `Function GT_PtrToString(ByVal ptr As Long) As Long` |

---

## crypto

| Hàm | Khai báo VBA |
|---|---|
| `SetKey` | `Function SetKey(ByVal key As String) As Long` |
| `Encrypt` | `Function Encrypt(ByVal plaintext As String) As LongPtr` |
| `Decrypt` | `Function Decrypt(ByVal ciphertext As String) As LongPtr` |

---

## datetime

| Hàm | Khai báo VBA |
|---|---|
| `DateTimeNow`<br><sub>==========================================</sub> | `Function DateTimeNow As LongPtr` |
| `DateNow` | `Function DateNow As LongPtr` |
| `TimeNow` | `Function TimeNow As LongPtr` |
| `DateTimeUTC` | `Function DateTimeUTC As LongPtr` |
| `UnixTimestamp` | `Function UnixTimestamp As LongLong` |
| `UnixTimestampMs` | `Function UnixTimestampMs As LongLong` |
| `DateTimeFormat`<br><sub>==========================================</sub> | `Function DateTimeFormat(ByVal dtStr As String, ByVal fmtIn As String, ByVal fmtOut As String) As LongPtr` |
| `UnixToDateTime` | `Function UnixToDateTime(ByVal unix As LongLong) As LongPtr` |
| `DateTimeToUnix` | `Function DateTimeToUnix(ByVal dtStr As String) As LongLong` |
| `DateTimeAddDays`<br><sub>==========================================</sub> | `Function DateTimeAddDays(ByVal dtStr As String, ByVal nDays As Long) As LongPtr` |
| `DateTimeAddMonths` | `Function DateTimeAddMonths(ByVal dtStr As String, ByVal nMonths As Long) As LongPtr` |
| `DateTimeAddYears` | `Function DateTimeAddYears(ByVal dtStr As String, ByVal nYears As Long) As LongPtr` |
| `DateTimeAddHours` | `Function DateTimeAddHours(ByVal dtStr As String, ByVal nHours As Long) As LongPtr` |
| `DateTimeAddMinutes` | `Function DateTimeAddMinutes(ByVal dtStr As String, ByVal nMins As Long) As LongPtr` |
| `DateTimeDiffDays`<br><sub>==========================================</sub> | `Function DateTimeDiffDays(ByVal dtStr1 As String, ByVal dtStr2 As String) As Long` |
| `DateTimeDiffHours` | `Function DateTimeDiffHours(ByVal dtStr1 As String, ByVal dtStr2 As String) As Long` |
| `DateTimeDiffMinutes` | `Function DateTimeDiffMinutes(ByVal dtStr1 As String, ByVal dtStr2 As String) As Long` |
| `DateTimeDiffSeconds` | `Function DateTimeDiffSeconds(ByVal dtStr1 As String, ByVal dtStr2 As String) As Long` |
| `DateTimeIsBefore` | `Function DateTimeIsBefore(ByVal dtStr1 As String, ByVal dtStr2 As String) As Long` |
| `DateTimeIsAfter` | `Function DateTimeIsAfter(ByVal dtStr1 As String, ByVal dtStr2 As String) As Long` |
| `DateGetYear`<br><sub>==========================================</sub> | `Function DateGetYear(ByVal dtStr As String) As Long` |
| `DateGetMonth` | `Function DateGetMonth(ByVal dtStr As String) As Long` |
| `DateGetDay` | `Function DateGetDay(ByVal dtStr As String) As Long` |
| `DateGetWeekday` | `Function DateGetWeekday(ByVal dtStr As String) As LongPtr` |
| `DateGetWeekNumber` | `Function DateGetWeekNumber(ByVal dtStr As String) As Long` |
| `DateGetDayOfYear` | `Function DateGetDayOfYear(ByVal dtStr As String) As Long` |
| `DateIsLeapYear` | `Function DateIsLeapYear(ByVal nYear As Long) As Long` |
| `DateGetDaysInMonth` | `Function DateGetDaysInMonth(ByVal nYear As Long, ByVal nMonth As Long) As Long` |
| `DateGetQuarter` | `Function DateGetQuarter(ByVal dtStr As String) As Long` |
| `DateTimeConvertTZ`<br><sub>==========================================</sub> | `Function DateTimeConvertTZ(ByVal dtStr As String, ByVal fromTZ As String, ByVal toTZ As String) As LongPtr` |

---

## encodeutils

| Hàm | Khai báo VBA |
|---|---|
| `Base64Encode`<br><sub>==========================================</sub> | `Function Base64Encode(ByVal s As String) As LongPtr` |
| `Base64Decode` | `Function Base64Decode(ByVal s As String) As LongPtr` |
| `Base64EncodeURL` | `Function Base64EncodeURL(ByVal s As String) As LongPtr` |
| `Base64DecodeURL` | `Function Base64DecodeURL(ByVal s As String) As LongPtr` |
| `Base64IsValid` | `Function Base64IsValid(ByVal s As String) As Long` |
| `HexEncode`<br><sub>==========================================</sub> | `Function HexEncode(ByVal s As String) As LongPtr` |
| `HexDecode` | `Function HexDecode(ByVal s As String) As LongPtr` |
| `HexIsValid` | `Function HexIsValid(ByVal s As String) As Long` |
| `HexToDecimal` | `Function HexToDecimal(ByVal s As String) As LongLong` |
| `DecimalToHex` | `Function DecimalToHex(ByVal n As LongLong) As LongPtr` |
| `DecimalToBinary` | `Function DecimalToBinary(ByVal n As LongLong) As LongPtr` |
| `BinaryToDecimal` | `Function BinaryToDecimal(ByVal s As String) As LongLong` |
| `DecimalToOctal` | `Function DecimalToOctal(ByVal n As LongLong) As LongPtr` |
| `OctalToDecimal` | `Function OctalToDecimal(ByVal s As String) As LongLong` |
| `URLEncode`<br><sub>==========================================</sub> | `Function URLEncode(ByVal s As String) As LongPtr` |
| `URLDecode` | `Function URLDecode(ByVal s As String) As LongPtr` |
| `URLPathEncode` | `Function URLPathEncode(ByVal s As String) As LongPtr` |
| `URLPathDecode` | `Function URLPathDecode(ByVal s As String) As LongPtr` |
| `HTMLEncode`<br><sub>==========================================</sub> | `Function HTMLEncode(ByVal s As String) As LongPtr` |
| `HTMLDecode` | `Function HTMLDecode(ByVal s As String) As LongPtr` |
| `UTF8IsValid`<br><sub>==========================================</sub> | `Function UTF8IsValid(ByVal s As String) As Long` |
| `UTF8RuneCount` | `Function UTF8RuneCount(ByVal s As String) As Long` |
| `UTF8ByteCount` | `Function UTF8ByteCount(ByVal s As String) As Long` |

---

## filecompression

| Hàm | Khai báo VBA |
|---|---|
| `GzipCompressFile`<br><sub>==========================================</sub> | `Function GzipCompressFile(ByVal srcPath As String, ByVal dstPath As String) As Long` |
| `GzipDecompressFile` | `Function GzipDecompressFile(ByVal srcPath As String, ByVal dstPath As String) As Long` |
| `ZlibCompressFile`<br><sub>==========================================</sub> | `Function ZlibCompressFile(ByVal srcPath As String, ByVal dstPath As String) As Long` |
| `ZlibDecompressFile` | `Function ZlibDecompressFile(ByVal srcPath As String, ByVal dstPath As String) As Long` |

---

## fileio

| Hàm | Khai báo VBA |
|---|---|
| `GT_FreeString`<br><sub>GT_FreeString: Giải phóng vùng nhớ mà Go đã cấp phát khi ...</sub> | `Sub GT_FreeString(ByVal s As String)` |
| `GT_GetFileSize`<br><sub>GT_GetFileSize: Lấy kích thước file (bytes). Thất bại trả...</sub> | `Function GT_GetFileSize(ByVal fPath As String) As LongLong` |
| `GT_FileExists`<br><sub>GT_FileExists: Kiểm tra file có tồn tại không. (1: Có, 0:...</sub> | `Function GT_FileExists(ByVal fPath As String) As Long` |
| `GT_FolderExists`<br><sub>GT_FolderExists: Kiểm tra thư mục có tồn tại không. (1: C...</sub> | `Function GT_FolderExists(ByVal fPath As String) As Long` |
| `GT_GetPathType`<br><sub>GT_GetPathType: Phân biệt đường dẫn. (1: File, 2: Folder,...</sub> | `Function GT_GetPathType(ByVal fPath As String) As Long` |
| `GT_ReadFile`<br><sub>GT_ReadFile: Đọc toàn bộ nội dung file trả về chuỗi.</sub> | `Function GT_ReadFile(ByVal fPath As String) As LongPtr` |
| `GT_WriteFile`<br><sub>GT_WriteFile: Ghi đè nội dung vào file. (1: Thành công, 0...</sub> | `Function GT_WriteFile(ByVal fPath As String, ByVal content As String) As Long` |
| `GT_AppendFile`<br><sub>GT_AppendFile: Ghi thêm vào cuối file. (1: Thành công, 0:...</sub> | `Function GT_AppendFile(ByVal fPath As String, ByVal content As String) As Long` |
| `GT_CopyFile`<br><sub>GT_CopyFile: Sao chép file. (1: Thành công, 0: Thất bại)</sub> | `Function GT_CopyFile(ByVal sSrc As String, ByVal sDst As String) As Long` |
| `GT_MoveFile`<br><sub>GT_MoveFile: Di chuyển file hoặc đổi tên.</sub> | `Function GT_MoveFile(ByVal sSrc As String, ByVal sDst As String) As Long` |
| `GT_DeleteFile`<br><sub>GT_DeleteFile: Xóa file.</sub> | `Function GT_DeleteFile(ByVal fPath As String) As Long` |
| `GT_CreateFolder`<br><sub>GT_CreateFolder: Tạo thư mục (bao gồm cả thư mục cha nếu ...</sub> | `Function GT_CreateFolder(ByVal fPath As String) As Long` |
| `GT_DeleteFolder`<br><sub>GT_DeleteFolder: Xóa thư mục và tất cả nội dung bên trong.</sub> | `Function GT_DeleteFolder(ByVal fPath As String) As Long` |
| `GT_ClearFolder`<br><sub>GT_ClearFolder: Xóa sạch nội dung trong thư mục nhưng giữ...</sub> | `Function GT_ClearFolder(ByVal fPath As String) As Long` |
| `GT_ListFiles`<br><sub>GT_ListFiles: Liệt kê file trong thư mục, cách nhau bằng ...</sub> | `Function GT_ListFiles(ByVal fPath As String) As LongPtr` |
| `GT_ListFolders`<br><sub>GT_ListFolders: Liệt kê thư mục con, cách nhau bằng dấu |</sub> | `Function GT_ListFolders(ByVal fPath As String) As LongPtr` |
| `GT_GetCurrentDir`<br><sub>GT_GetCurrentDir: Lấy thư mục đang chạy ứng dụng.</sub> | `Function GT_GetCurrentDir As LongPtr` |
| `GT_GetTempDir`<br><sub>GT_GetTempDir: Lấy đường dẫn thư mục Temp của Windows.</sub> | `Function GT_GetTempDir As LongPtr` |
| `GT_GetFileModTime`<br><sub>GT_GetFileModTime: Lấy ngày giờ cập nhật cuối (Định dạng:...</sub> | `Function GT_GetFileModTime(ByVal fPath As String) As LongPtr` |
| `GT_GetFileExtension`<br><sub>GT_GetFileExtension: Lấy phần mở rộng (ví dụ: .exe, .txt)</sub> | `Function GT_GetFileExtension(ByVal fPath As String) As LongPtr` |
| `GT_GetFileName`<br><sub>GT_GetFileName: Lấy tên file kèm đuôi từ đường dẫn đầy đủ.</sub> | `Function GT_GetFileName(ByVal fPath As String) As LongPtr` |

---

## hash

| Hàm | Khai báo VBA |
|---|---|
| `MD5Hash` | `Function MD5Hash(ByVal input As String) As LongPtr` |
| `SHA1Hash` | `Function SHA1Hash(ByVal input As String) As LongPtr` |
| `SHA256Hash` | `Function SHA256Hash(ByVal input As String) As LongPtr` |
| `SHA512Hash` | `Function SHA512Hash(ByVal input As String) As LongPtr` |

---

## http

| Hàm | Khai báo VBA |
|---|---|
| `HttpSetTimeout`<br><sub>==========================================</sub> | `Sub HttpSetTimeout(ByVal seconds As Long)` |
| `HttpGet`<br><sub>==========================================</sub> | `Function HttpGet(ByVal reqUrl As String) As LongPtr` |
| `HttpGetWithHeader` | `Function HttpGetWithHeader(ByVal reqUrl As String, ByVal headerKey As String, ByVal headerVal As String) As LongPtr` |
| `HttpPostJSON`<br><sub>==========================================</sub> | `Function HttpPostJSON(ByVal reqUrl As String, ByVal jsonBody As String) As LongPtr` |
| `HttpPostForm` | `Function HttpPostForm(ByVal reqUrl As String, ByVal formData As String) As LongPtr` |
| `HttpPut`<br><sub>==========================================</sub> | `Function HttpPut(ByVal reqUrl As String, ByVal jsonBody As String) As LongPtr` |
| `HttpPatch` | `Function HttpPatch(ByVal reqUrl As String, ByVal jsonBody As String) As LongPtr` |
| `HttpDelete` | `Function HttpDelete(ByVal reqUrl As String) As LongPtr` |
| `HttpGetStatusCode`<br><sub>==========================================</sub> | `Function HttpGetStatusCode(ByVal reqUrl As String) As Long` |
| `HttpDownloadFile`<br><sub>==========================================</sub> | `Function HttpDownloadFile(ByVal reqUrl As String, ByVal dstPath As String) As Long` |
| `JsonGetValue`<br><sub>==========================================</sub> | `Function JsonGetValue(ByVal jsonStr As String, ByVal key As String) As LongPtr` |
| `JsonBuildObject` | `Function JsonBuildObject(ByVal kvPairs As String) As LongPtr` |

---

## jwt

| Hàm | Khai báo VBA |
|---|---|
| `JwtGenerateKeys`<br><sub>==========================================</sub> | `Function JwtGenerateKeys(ByVal privatePath As String, ByVal publicPath As String) As Long` |
| `JwtLoadPrivateKey` | `Function JwtLoadPrivateKey(ByVal fPath As String) As Long` |
| `JwtLoadPublicKey` | `Function JwtLoadPublicKey(ByVal fPath As String) As Long` |
| `JwtCreateToken`<br><sub>==========================================</sub> | `Function JwtCreateToken(ByVal sub As String, ByVal role As String, ByVal iss As String, ByVal aud As String, ByVal expSeconds As Long) As LongPtr` |
| `JwtVerifyToken`<br><sub>==========================================</sub> | `Function JwtVerifyToken(ByVal token As String) As Long` |
| `JwtGetClaims`<br><sub>==========================================</sub> | `Function JwtGetClaims(ByVal token As String) As LongPtr` |
| `JwtGetField` | `Function JwtGetField(ByVal token As String, ByVal field As String) As LongPtr` |
| `JwtIsExpired` | `Function JwtIsExpired(ByVal token As String) As Long` |
| `JwtGetExpireTime` | `Function JwtGetExpireTime(ByVal token As String) As LongPtr` |
| `JwtRefreshToken`<br><sub>==========================================</sub> | `Function JwtRefreshToken(ByVal token As String, ByVal expSeconds As Long) As LongPtr` |

---

## stringutils

| Hàm | Khai báo VBA |
|---|---|
| `StrToUpper` | `Function StrToUpper(ByVal s As String) As LongPtr` |
| `StrToLower` | `Function StrToLower(ByVal s As String) As LongPtr` |
| `StrTrim` | `Function StrTrim(ByVal s As String) As LongPtr` |
| `StrReplace` | `Function StrReplace(ByVal s As String, ByVal old As String, ByVal new As String) As LongPtr` |
| `StrContains` | `Function StrContains(ByVal s As String, ByVal sub As String) As Long` |
| `StrStartsWith` | `Function StrStartsWith(ByVal s As String, ByVal prefix As String) As Long` |
| `StrEndsWith` | `Function StrEndsWith(ByVal s As String, ByVal suffix As String) As Long` |
| `StrLength` | `Function StrLength(ByVal s As String) As Long` |
| `StrReverse` | `Function StrReverse(ByVal s As String) As LongPtr` |
| `StrCount` | `Function StrCount(ByVal s As String, ByVal sub As String) As Long` |
| `StrIsNumeric` | `Function StrIsNumeric(ByVal s As String) As Long` |
| `StrIsPalindrome` | `Function StrIsPalindrome(ByVal s As String) As Long` |

---

## validation

| Hàm | Khai báo VBA |
|---|---|
| `ValidateEmail`<br><sub>==========================================</sub> | `Function ValidateEmail(ByVal s As String) As Long` |
| `ValidatePhone`<br><sub>==========================================</sub> | `Function ValidatePhone(ByVal s As String) As Long` |
| `ValidatePhoneVN` | `Function ValidatePhoneVN(ByVal s As String) As Long` |
| `ValidateURL`<br><sub>==========================================</sub> | `Function ValidateURL(ByVal s As String) As Long` |
| `ValidateIPv4`<br><sub>==========================================</sub> | `Function ValidateIPv4(ByVal s As String) As Long` |
| `ValidateIPv6` | `Function ValidateIPv6(ByVal s As String) As Long` |
| `ValidateIP` | `Function ValidateIP(ByVal s As String) As Long` |
| `ValidateInteger`<br><sub>==========================================</sub> | `Function ValidateInteger(ByVal s As String) As Long` |
| `ValidateFloat` | `Function ValidateFloat(ByVal s As String) As Long` |
| `ValidateInRange` | `Function ValidateInRange(ByVal s As String, ByVal nMin As Double, ByVal nMax As Double) As Long` |
| `ValidateMinLength`<br><sub>==========================================</sub> | `Function ValidateMinLength(ByVal s As String, ByVal nMin As Long) As Long` |
| `ValidateMaxLength` | `Function ValidateMaxLength(ByVal s As String, ByVal nMax As Long) As Long` |
| `ValidateNotEmpty` | `Function ValidateNotEmpty(ByVal s As String) As Long` |
| `ValidateAlpha` | `Function ValidateAlpha(ByVal s As String) As Long` |
| `ValidateAlphaNumeric` | `Function ValidateAlphaNumeric(ByVal s As String) As Long` |
| `ValidateContainsUpper` | `Function ValidateContainsUpper(ByVal s As String) As Long` |
| `ValidateContainsLower` | `Function ValidateContainsLower(ByVal s As String) As Long` |
| `ValidateContainsDigit` | `Function ValidateContainsDigit(ByVal s As String) As Long` |
| `ValidateContainsSpecial` | `Function ValidateContainsSpecial(ByVal s As String) As Long` |
| `ValidatePassword`<br><sub>==========================================</sub> | `Function ValidatePassword(ByVal s As String) As Long` |
| `ValidatePasswordStrength` | `Function ValidatePasswordStrength(ByVal s As String) As Long` |
| `ValidateDate`<br><sub>==========================================</sub> | `Function ValidateDate(ByVal s As String, ByVal fmt As String) As Long` |
| `ValidateCreditCard`<br><sub>==========================================</sub> | `Function ValidateCreditCard(ByVal s As String) As Long` |
| `ValidateRegex`<br><sub>==========================================</sub> | `Function ValidateRegex(ByVal s As String, ByVal pattern As String) As Long` |

---

## zipfile

| Hàm | Khai báo VBA |
|---|---|
| `ZipCompressFile` | `Function ZipCompressFile(ByVal srcPath As String, ByVal dstZip As String) As Long` |
| `ZipCompressFolder` | `Function ZipCompressFolder(ByVal srcFolder As String, ByVal dstZip As String) As Long` |
| `ZipDecompress` | `Function ZipDecompress(ByVal srcZip As String, ByVal dstFolder As String) As Long` |
| `ZipAddFile` | `Function ZipAddFile(ByVal existingZip As String, ByVal newFile As String) As Long` |
| `ZipListFiles` | `Function ZipListFiles(ByVal srcZip As String) As LongPtr` |
| `ZipGetFileCount` | `Function ZipGetFileCount(ByVal srcZip As String) As Long` |

---

_Sinh bởi **GoDocGen** `Mode=vba` — 05/05/2026 07:48_
