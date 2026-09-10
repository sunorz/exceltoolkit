/**
 * ============================================================================
 * excel_toolkit (main.cpp) — Excel 暗水印与保护工具（C++ 独立命令行版）
 * ============================================================================
 */

#define _CRT_SECURE_NO_WARNINGS
#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>
#include <algorithm>
#include <chrono>
#include <iomanip>
#include <filesystem>
#include <cstdlib>

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <conio.h>
#endif

#include "excel_toolkit_core.h"

namespace fs = std::filesystem;

std::string GetAppDir() {
#ifdef _WIN32
    wchar_t szPath[MAX_PATH];
    if (GetModuleFileNameW(NULL, szPath, MAX_PATH)) {
        std::wstring wpath(szPath);
        size_t lastSlash = wpath.find_last_of(L"\\/");
        if (lastSlash != std::wstring::npos) {
            wpath = wpath.substr(0, lastSlash);
        }
        int sizeNeeded = WideCharToMultiByte(CP_UTF8, 0, wpath.c_str(), -1, NULL, 0, NULL, NULL);
        std::string path(sizeNeeded, 0);
        WideCharToMultiByte(CP_UTF8, 0, wpath.c_str(), -1, &path[0], sizeNeeded, NULL, NULL);
        if (!path.empty() && path.back() == '\0') path.pop_back();
        return path;
    }
#endif
    return fs::current_path().u8string();
}

std::vector<std::string> ScanTargetFiles(const std::string& scriptDir) {
    std::vector<std::string> files;
    try {
        for (const auto& entry : fs::directory_iterator(fs::u8path(scriptDir))) {
            if (entry.is_regular_file()) {
                std::string filename = entry.path().filename().u8string();
                std::string ext = entry.path().extension().u8string();
                std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
                if (ext == ".xlsx" || ext == ".xls") {
                    files.push_back(filename);
                }
            }
        }
        std::sort(files.begin(), files.end());
    } catch (...) {}
    return files;
}

int main(int argc, char* argv[]) {
#ifdef _WIN32
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
#endif

    std::cout << "=======================================================\n";
    std::cout << "  excel-toolkit — Excel 暗水印与保护工具 (C++ 原生版)\n";
    std::cout << "=======================================================\n";
    
    if (argc < 2) {
        std::cout << "用法: \n";
        std::cout << "  excel-toolkit <excel文件路径> [选项]\n\n";
        std::cout << "选项:\n";
        std::cout << "  --add-watermark <文字>   写入 veryHidden 暗水印\n";
        std::cout << "  --trace-id <溯源编号>    指定防伪溯源编号\n";
        std::cout << "  --lock-structure         锁定工作簿结构\n";
        std::cout << "  --protect-sheet [密码]   锁定工作表内容\n";
        std::cout << "  --unprotect              一键解除所有保护\n";
        std::cout << "  --inspect                检查提取暗水印与结构全景\n";
        return 0;
    }

    std::string filePath = argv[1];
    std::string mode = (argc >= 3) ? argv[2] : "--inspect";

    if (mode == "--inspect") {
        ExcelToolkitCore::InspectWatermarkInfo info;
        std::string err;
        if (ExcelToolkitCore::InspectWatermarkFile(filePath, info, err)) {
            std::cout << "[成功] 解析完成:\n";
            std::cout << "  工作簿结构锁定: " << (info.isWorkbookLocked ? "是" : "否") << "\n";
            std::cout << "  包含暗水印表: " << (info.hasHiddenSheet ? "是" : "否") << "\n";
            if (!info.traceId.empty()) {
                std::cout << "  溯源编号: " << info.traceId << "\n";
                std::cout << "  水印文字: " << info.watermarkText << "\n";
                std::cout << "  写入时间: " << info.createdAt << "\n";
            }
        } else {
            std::cerr << "[错误] " << err << "\n";
        }
    }

    return 0;
}
