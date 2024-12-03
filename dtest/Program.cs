using System;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;

class Program {
    static void Main() {
        string inputFilePath = @"C:\Users\hwcho\OneDrive\바탕 화면\weekly_log_automation\test.hwp";  // 원본 HWP 파일 경로
        string outputFilePath = @"C:\Users\hwcho\OneDrive\바탕 화면\weekly_log_automation\test.hwpx"; // 변환될 HWPX 파일 경로
        string txtOutputFilePath = @"C:\Users\hwcho\OneDrive\바탕 화면\weekly_log_automation\test.txt"; // 저장될 TXT 파일 경로

        // 1. HWP 파일을 HWPX 형식으로 변환하기
        ConvertHwpToHwpx(inputFilePath, outputFilePath);

        // 2. 변환된 HWPX 파일에서 텍스트 추출 및 TXT 파일로 저장하기
        ExtractTextFromHwpxAndSave(outputFilePath, txtOutputFilePath);
    }

    static void ConvertHwpToHwpx(string inputFilePath, string outputFilePath) {
        // 한글 HwpCtrl 객체 초기화
        dynamic hwp = Activator.CreateInstance(Type.GetTypeFromProgID("HWPFrame.HwpObject"));

        if (hwp != null) {
            try {
                // 한글 숨김 실행 설정
                hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule");

                // HWP 파일 열기
                if (hwp.Open(inputFilePath, "", "")) {
                    Console.WriteLine("파일을 성공적으로 열었습니다.");

                    // HWPX 파일로 저장
                    bool isSaved = hwp.SaveAs(outputFilePath, "HWPX", "");

                    if (isSaved) {
                        Console.WriteLine("파일을 성공적으로 HWPX로 저장했습니다.");
                    } else {
                        Console.WriteLine("파일을 저장하는 데 실패했습니다.");
                    }

                    // 문서 닫기
                    hwp.Clear(1);
                } else {
                    Console.WriteLine("파일을 여는 데 실패했습니다.");
                }
            } finally {
                // 한글 객체 해제
                Marshal.ReleaseComObject(hwp);
                hwp = null;
            }
        } else {
            Console.WriteLine("HwpObject를 생성하는 데 실패했습니다. 한컴오피스가 설치되어 있는지 확인하십시오.");
        }
    }

    static void ExtractTextFromHwpxAndSave(string hwpxFilePath, string txtOutputFilePath) {
        try {
            // HWPX 파일을 ZIP으로 해제
            using (ZipArchive archive = ZipFile.OpenRead(hwpxFilePath)) {
                using (StreamWriter writer = new StreamWriter(txtOutputFilePath)) {
                    foreach (ZipArchiveEntry entry in archive.Entries) {
                        // XML 파일만 처리
                        if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) {
                            Console.WriteLine($"XML 파일 처리 중: {entry.FullName}");
                            using (Stream stream = entry.Open()) {
                                try {
                                    XDocument xmlDoc = XDocument.Load(stream);

                                    // 모든 텍스트 노드를 찾아내기
                                    var textNodes = xmlDoc.Descendants();

                                    foreach (var node in textNodes) {
                                        if (!string.IsNullOrWhiteSpace(node.Value)) {
                                            string textContent = node.Value.Trim();

                                            // 텍스트 파일에 저장
                                            writer.WriteLine(textContent);
                                        }
                                    }
                                } catch (XmlException ex) {
                                    Console.WriteLine($"XML 파일 로드 중 오류 발생: {ex.Message}");
                                    // XML 로드 실패 시 무시하고 다음 파일 처리
                                    continue;
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine("텍스트를 성공적으로 TXT 파일로 저장했습니다.");
        } catch (Exception ex) {
            Console.WriteLine("오류 발생: " + ex.Message);
        }
    }
}
