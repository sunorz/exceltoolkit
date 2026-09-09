/**
 * ============================================================================
 * excel_toolkit (main.cpp) — Excel 暗水印与保护工具（C++ 独立三合一）
 * 
 * 功能：
 *   1. 添加暗水印   —— 在文件里藏一个 veryHidden 隐藏工作表，写入水印文字/溯源编号/时间
 *   2. 查看暗水印   —— 列出文件的 sheet 及可见性，读取隐藏表里的溯源信息
 *   3. 保护内容（禁止编辑）—— 给工作表加密码保护，普通用户无法直接编辑单元格；
 *                       同时锁定工作簿结构，无法通过菜单增删/显示隐藏工作表
 *                       （这样也能防止别人从"取消隐藏"菜单里直接看到暗水印表）
 * 
 * 依赖：纯 C++17 标准库与嵌入式 miniz，零额外动态库，直接双击运行
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

// 获取程序所在物理目录
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

// 扫描当前目录下的 .xlsx 与 .xls
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

// 交互式选择文件
std::vector<std::string> ChooseFilesInteractively(const std::vector<std::string>& files, const std::string& promptSuffix = "") {
    std::cout << "在当前文件夹找到以下文件：\n\n";
    for (size_t i = 0; i < files.size(); ++i) {
        std::cout << "  " << (i + 1) << ". " << files[i] << "\n";
    }
    std::cout << "\n共 " << files.size() << " 个文件。\n";
    std::cout << "请输入要处理的文件序号（例如: 1,3,5），或直接回车表示【全部处理】" << promptSuffix << "：\n> ";

    std::string raw;
    std::getline(std::cin, raw);

    raw.erase(0, raw.find_first_not_of(" \t\r\n"));
    raw.erase(raw.find_last_not_of(" \t\r\n") + 1);

    if (raw.empty()) {
        return files;
    }

    for (size_t i = 0; i < raw.size(); ++i) {
        if ((unsigned char)raw[i] == 0xef && i + 2 < raw.size() &&
            (unsigned char)raw[i+1] == 0xbc && (unsigned char)raw[i+2] == 0x8c) {
            raw[i] = ','; raw[i+1] = ' '; raw[i+2] = ' ';
        }
    }

    std::stringstream ss(raw);
    std::string item;
    std::vector<int> indices;
    while (std::getline(ss, item, ',')) {
        item.erase(0, item.find_first_not_of(" \t\r\n"));
        item.erase(item.find_last_not_of(" \t\r\n") + 1);
        if (!item.empty()) {
            try {
                int idx = std::stoi(item);
                if (idx >= 1 && idx <= (int)files.size()) {
                    indices.push_back(idx);
                }
            } catch (...) {}
        }
    }

    if (indices.empty()) {
        return files;
    }

    std::sort(indices.begin(), indices.end());
    indices.erase(std::unique(indices.begin(), indices.end()), indices.end());

    std::vector<std::string> selected;
    for (int idx : indices) {
        selected.push_back(files[idx - 1]);
    }
    return selected;
}

// 检查并过滤 .xls
std::vector<std::string> WarnAndFilterXls(const std::vector<std::string>& targets) {
    bool hasXls = false;
    for (const auto& f : targets) {
        if (f.size() >= 4 && f.substr(f.size() - 4) == ".xls") {
            hasXls = true;
            break;
        }
    }
    if (!hasXls) return targets;

    if (ExcelToolkitCore::CheckSofficeAvailable()) {
        return targets;
    }

    std::cout << "\n=======================================================\n";
    std::cout << "  无法处理 .xls 文件\n";
    std::cout << "=======================================================\n";
    std::cout << "检测到你选择的文件里有 .xls（旧版 Excel 97-2003）文件，\n"
              << "但处理 .xls 需要依赖 LibreOffice 做格式转换，当前电脑上\n"
              << "没有找到 LibreOffice（或未加入系统 PATH）。\n\n"
              << "请选择以下任一方式后重新运行：\n"
              << "  1. 用 Excel/WPS 打开该文件，另存为 .xlsx 格式，再运行本工具；\n"
              << "  2. 或安装 LibreOffice（https://www.libreoffice.org/download），\n"
              << "     安装后重新运行即可自动处理 .xls。\n";
    std::cout << "=======================================================\n\n";

    std::vector<std::string> filtered;
    for (const auto& f : targets) {
        if (!(f.size() >= 4 && f.substr(f.size() - 4) == ".xls")) {
            filtered.push_back(f);
        }
    }
    return filtered;
}

// 密码交互录入
std::string AskPassword(bool confirm = true) {
    while (true) {
        std::cout << "请输入保护密码（输入时不会显示字符）：";
        std::string pw;
#ifdef _WIN32
        char ch;
        while ((ch = _getch()) != '\r' && ch != '\n') {
            if (ch == '\b') {
                if (!pw.empty()) pw.pop_back();
            } else if (ch >= 32 && ch <= 126) {
                pw.push_back(ch);
            }
        }
        std::cout << "\n";
#else
        std::getline(std::cin, pw);
#endif
        if (pw.empty()) {
            std::cout << "密码不能为空，请重新输入。\n";
            continue;
        }
        if (!confirm) return pw;

        std::cout << "请再输入一次确认：";
        std::string pw2;
#ifdef _WIN32
        while ((ch = _getch()) != '\r' && ch != '\n') {
            if (ch == '\b') {
                if (!pw2.empty()) pw2.pop_back();
            } else if (ch >= 32 && ch <= 126) {
                pw2.push_back(ch);
            }
        }
        std::cout << "\n";
#else
        std::getline(std::cin, pw2);
#endif
        if (pw != pw2) {
            std::cout << "两次输入的密码不一致，请重新输入。\n";
            continue;
        }
        return pw;
    }
}

// 功能 1：添加暗水印
void ActionAddWatermark(const std::string& scriptDir) {
    std::cout << "\n=======================================================\n";
    std::cout << "  功能 1：添加暗水印\n";
    std::cout << "=======================================================\n\n";

    auto files = ScanTargetFiles(scriptDir);
    if (files.empty()) {
        std::cout << "未在 " << scriptDir << " 找到 .xlsx / .xls 文件。\n";
        return;
    }

    auto selected = ChooseFilesInteractively(files);
    if (selected.empty()) {
        std::cout << "未选择任何文件。\n";
        return;
    }

    std::cout << "\n";
    std::string watermarkText;
    while (true) {
        std::cout << "请输入要写入的水印文字（例如: 给交通局的表）：";
        std::getline(std::cin, watermarkText);
        watermarkText.erase(0, watermarkText.find_first_not_of(" \t\r\n"));
        watermarkText.erase(watermarkText.find_last_not_of(" \t\r\n") + 1);
        if (!watermarkText.empty()) break;
        std::cout << "水印文字不能为空，请重新输入。\n";
    }

    std::string traceIdInput;
    if (selected.size() == 1) {
        std::cout << "请输入溯源编号（可直接回车，自动生成随机编号）：";
        std::getline(std::cin, traceIdInput);
        traceIdInput.erase(0, traceIdInput.find_first_not_of(" \t\r\n"));
        traceIdInput.erase(traceIdInput.find_last_not_of(" \t\r\n") + 1);
    } else {
        std::cout << "即将处理 " << selected.size() << " 个文件。\n"
                  << "  - 直接回车：每个文件自动生成不同的随机编号（推荐，方便区分是哪份文件流出）\n"
                  << "  - 输入内容：所有文件使用同一个编号\n"
                  << "请输入溯源编号，或直接回车：";
        std::getline(std::cin, traceIdInput);
        traceIdInput.erase(0, traceIdInput.find_first_not_of(" \t\r\n"));
        traceIdInput.erase(traceIdInput.find_last_not_of(" \t\r\n") + 1);
    }

    auto validTargets = WarnAndFilterXls(selected);
    if (validTargets.empty()) return;

    std::cout << "\n开始处理 " << validTargets.size() << " 个文件，水印文字: \"" << watermarkText << "\"\n\n";

    for (const auto& name : validTargets) {
        std::string fullPath = (fs::u8path(scriptDir) / fs::u8path(name)).u8string();
        std::string traceId = traceIdInput.empty() ? ExcelToolkitCore::GenerateRandomHex8() : traceIdInput;
        std::string err;
        
        bool ok = ExcelToolkitCore::EmbedWatermark(fullPath, watermarkText, traceId, err);
        if (ok) {
            std::cout << "[完成] " << name << " | 溯源编号: " << traceId << "\n";
        } else {
            std::cout << "[失败] " << name << ": " << err << "\n";
        }
    }

    std::cout << "\n全部处理完毕。建议：把每个文件对应的溯源编号记录下来\n"
              << "（比如发给谁对应哪个编号），以后用【功能2 查看暗水印】核对来源。\n";
}

// 功能 2：查看暗水印
void ActionViewWatermark(const std::string& scriptDir) {
    std::cout << "\n=======================================================\n";
    std::cout << "  功能 2：查看暗水印\n";
    std::cout << "=======================================================\n\n";

    auto files = ScanTargetFiles(scriptDir);
    if (files.empty()) {
        std::cout << "未在 " << scriptDir << " 找到 .xlsx / .xls 文件。\n";
        return;
    }

    std::cout << "找到以下文件：\n";
    for (size_t i = 0; i < files.size(); ++i) {
        std::cout << "  " << (i + 1) << ". " << files[i] << "\n";
    }

    int choiceIdx = -1;
    while (true) {
        std::cout << "\n请输入序号选择要检查的文件 (1-" << files.size() << ")，或输入 q 返回菜单：";
        std::string choice;
        std::getline(std::cin, choice);
        choice.erase(0, choice.find_first_not_of(" \t\r\n"));
        choice.erase(choice.find_last_not_of(" \t\r\n") + 1);
        if (choice == "q" || choice == "Q") return;
        try {
            int num = std::stoi(choice);
            if (num >= 1 && num <= (int)files.size()) {
                choiceIdx = num - 1;
                break;
            }
        } catch (...) {}
        std::cout << "输入无效，请重新输入。\n";
    }

    std::string targetName = files[choiceIdx];
    std::string fullPath = (fs::u8path(scriptDir) / fs::u8path(targetName)).u8string();
    std::cout << "\n正在检查：" << targetName << " ...\n";

    ExcelToolkitCore::InspectWatermarkInfo info;
    std::string err;
    if (!ExcelToolkitCore::InspectWatermarkFile(fullPath, info, err)) {
        std::cout << "读取失败: " << err << "\n";
        return;
    }

    std::cout << "\n工作表列表（括号内为可见性状态）：\n";
    for (const auto& ws : info.sheets) {
        std::cout << "  - " << ws.name << " (" << ws.state << ")\n";
    }

    std::cout << "\n=======================================================\n";
    std::cout << "  溯源信息\n";
    std::cout << "=======================================================\n";

    if (!info.hasHiddenSheet) {
        std::cout << "  未检测到本工具写入的暗水印（没有找到隐藏表 " << HIDDEN_SHEET_NAME_CONST << "）\n";
        std::cout << "=======================================================\n";
        return;
    }

    std::cout << "  溯源编号: " << (info.traceId.empty() ? "UNSPECIFIED" : info.traceId) << "\n";
    std::cout << "  水印文字: " << info.watermarkText << "\n";
    std::cout << "  生成时间: " << info.createdAt << "\n";
    std::cout << "=======================================================\n";
}

// 功能 3：保护内容（禁止编辑 / 解除保护）
void ActionProtect(const std::string& scriptDir) {
    std::cout << "\n=======================================================\n";
    std::cout << "  功能 3：保护内容（禁止编辑）\n";
    std::cout << "=======================================================\n";
    std::cout << "\n说明：这是 Excel 原生的工作表/工作簿保护，效果等同于\n"
              << "「审阅 -> 保护工作表」，没有密码无法编辑单元格、增删/显示隐藏工作表。\n"
              << "⚠️ 这不是加密，只是限制普通操作，专业工具几秒钟能破解密码，\n"
              << "适合防止同事手滑改错数据，不适合防止蓄意泄露/篡改。\n\n";

    std::cout << "请选择：1=开启保护  2=解除保护  q=返回菜单\n> ";
    std::string sub;
    std::getline(std::cin, sub);
    sub.erase(0, sub.find_first_not_of(" \t\r\n"));
    sub.erase(sub.find_last_not_of(" \t\r\n") + 1);

    if (sub == "q" || sub == "Q") return;
    if (sub != "1" && sub != "2") {
        std::cout << "输入无效，返回菜单。\n";
        return;
    }

    auto files = ScanTargetFiles(scriptDir);
    if (files.empty()) {
        std::cout << "未在 " << scriptDir << " 找到 .xlsx / .xls 文件。\n";
        return;
    }

    auto selected = ChooseFilesInteractively(files);
    if (selected.empty()) {
        std::cout << "未选择任何文件。\n";
        return;
    }

    auto validTargets = WarnAndFilterXls(selected);
    if (validTargets.empty()) return;

    std::cout << "\n";
    std::string password;
    if (sub == "1") {
        password = AskPassword(true);
    } else {
        password = AskPassword(false);
        std::cout << "提示：本工具解除保护将直接清空工作簿与工作表的保护设置。\n"
                  << "这里要求输入密码只是走个流程、方便你自己记录操作，不是真正校验。\n";
    }

    std::cout << "\n开始处理 " << validTargets.size() << " 个文件...\n\n";

    for (const auto& name : validTargets) {
        std::string fullPath = (fs::u8path(scriptDir) / fs::u8path(name)).u8string();
        std::string err;
        bool ok = false;
        if (sub == "1") {
            ok = ExcelToolkitCore::ProtectWorkbook(fullPath, password, err);
            if (ok) {
                std::cout << "[完成] " << name << " 已开启保护\n";
            } else {
                std::cout << "[失败] " << name << ": " << err << "\n";
            }
        } else {
            ok = ExcelToolkitCore::UnprotectWorkbook(fullPath, err);
            if (ok) {
                std::cout << "[完成] " << name << " 已解除保护\n";
            } else {
                std::cout << "[失败] " << name << ": " << err << "\n";
            }
        }
    }

    std::cout << "\n全部处理完毕。\n";
}

int main() {
#ifdef _WIN32
    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
#endif

    std::string scriptDir = GetAppDir();

    while (true) {
        std::cout << "\n=======================================================\n";
        std::cout << "  Excel 暗水印与保护工具 (Excel Toolkit)\n";
        std::cout << "=======================================================\n";
        std::cout << "\n请选择功能：\n"
                  << "  1. 添加暗水印\n"
                  << "  2. 查看暗水印\n"
                  << "  3. 保护内容（禁止编辑 / 解除保护）\n"
                  << "  q. 退出\n\n> ";

        std::string choice;
        std::getline(std::cin, choice);
        choice.erase(0, choice.find_first_not_of(" \t\r\n"));
        choice.erase(choice.find_last_not_of(" \t\r\n") + 1);

        if (choice == "1") {
            ActionAddWatermark(scriptDir);
        } else if (choice == "2") {
            ActionViewWatermark(scriptDir);
        } else if (choice == "3") {
            ActionProtect(scriptDir);
        } else if (choice == "q" || choice == "Q") {
            break;
        } else {
            std::cout << "输入无效，请重新选择。\n";
            continue;
        }

        std::cout << "\n按回车键返回主菜单...";
        std::string dummy;
        std::getline(std::cin, dummy);
    }

    return 0;
}
