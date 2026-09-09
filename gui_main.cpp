/**
 * ============================================================================
 * excel-toolkit (gui_main.cpp) — Windows 原生独立图形界面版 (GUI)
 * 
 * 核心设计：
 *   - 纯原生 Win32 + Common Controls v6 界面，无黑框控制台 (/SUBSYSTEM:WINDOWS)
 *   - 支持全窗口拖拽 Excel (.xlsx / .xls) 文件
 *   - 模块化复选框 (Checkbox) 组合设计：
 *       [✓] 添加暗水印 (写入 veryHidden 隐藏表)
 *       [✓] 锁定工作簿结构 (防止通过 Excel 菜单取消隐藏)
 *       [ ] 锁定工作表内容 (只读防编辑)
 *       [ ] 一键解除所有保护
 *   - 内存级 OpenXML 解析，杜绝中文路径与文件锁死导致的闪退
 *   - 全面 Exception Guard 异常拦截，100% 杜绝程序崩溃
 * ============================================================================
 */

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <string>
#include <vector>
#include <sstream>
#include <iomanip>
#include <chrono>
#include "excel_toolkit_core.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "comdlg32.lib")

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

// 控件 ID
#define IDC_TABCONTROL           1001
#define IDC_EDIT_FILEPATH        1002
#define IDC_BTN_BROWSE           1003

// Tab 1 (综合处理)
#define IDC_CHK_ADD_WM           1004
#define IDC_EDIT_WM_TEXT         1005
#define IDC_EDIT_TRACE_ID        1006
#define IDC_BTN_GEN_TRACE        1007
#define IDC_CHK_LOCK_STRUCT      1008
#define IDC_CHK_PROTECT_SHEET    1009
#define IDC_EDIT_PASSWORD        1010
#define IDC_CHK_UNPROTECT_ALL    1011
#define IDC_BTN_EXECUTE_PROCESS  1012

// Tab 2 (查看取证)
#define IDC_BTN_START_INSPECT    1013
#define IDC_LIST_SHEETS          1014
#define IDC_EDIT_INSPECT_RESULT  1015

// 公共
#define IDC_STATIC_STATUS        1020

// 全局变量
HINSTANCE g_hInstance = NULL;
HWND g_hMainWnd = NULL;
HWND g_hTabCtrl = NULL;
HWND g_hEditFilePath = NULL;

HWND g_hChkAddWm = NULL;
HWND g_hLblWmText = NULL;
HWND g_hEditWmText = NULL;
HWND g_hLblTraceId = NULL;
HWND g_hEditTraceId = NULL;
HWND g_hBtnGenTrace = NULL;
HWND g_hChkLockStruct = NULL;
HWND g_hChkProtectSheet = NULL;
HWND g_hLblPassword = NULL;
HWND g_hEditPassword = NULL;
HWND g_hChkUnprotectAll = NULL;
HWND g_hBtnExecuteProcess = NULL;

HWND g_hBtnStartInspect = NULL;
HWND g_hLblSheetsList = NULL;
HWND g_hListSheets = NULL;
HWND g_hLblInspectRes = NULL;
HWND g_hEditInspectResult = NULL;

HWND g_hStaticStatus = NULL;
HFONT g_hFont = NULL;

int g_currentTab = 0;
std::wstring g_selectedFilePath = L"";

std::wstring Utf8ToWString(const std::string& str) {
    if (str.empty()) return L"";
    int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    if (size <= 1) return L"";
    std::wstring result(size - 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, &result[0], size);
    return result;
}

std::string WStringToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return "";
    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, NULL, 0, NULL, NULL);
    if (size <= 1) return "";
    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, NULL, NULL);
    return result;
}

std::wstring GenerateTraceId() {
    auto now = std::chrono::system_clock::now();
    auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(now.time_since_epoch()).count();
    std::wstringstream wss;
    wss << L"SEC-" << std::hex << std::uppercase << (ms & 0xFFFFFF) << L"-" << (rand() % 9000 + 1000);
    return wss.str();
}

void OnBrowseFile() {
    try {
        wchar_t szFile[MAX_PATH] = { 0 };
        OPENFILENAMEW ofn = { 0 };
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = g_hMainWnd;
        ofn.lpstrFilter = L"Excel 工作簿 (*.xlsx;*.xls)\0*.xlsx;*.xls\0所有文件 (*.*)\0*.*\0";
        ofn.lpstrFile = szFile;
        ofn.nMaxFile = MAX_PATH;
        ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;

        if (GetOpenFileNameW(&ofn)) {
            g_selectedFilePath = szFile;
            SetWindowTextW(g_hEditFilePath, szFile);
            SetWindowTextW(g_hStaticStatus, L"就绪：已选择文件。");
            if (g_currentTab == 1) {
                SendMessage(g_hMainWnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_START_INSPECT, BN_CLICKED), 0);
            }
        }
    } catch (...) {}
}

void OnExecuteProcess() {
    if (g_selectedFilePath.empty()) {
        MessageBoxW(g_hMainWnd, L"请先选择或拖入待处理的 Excel (.xlsx / .xls) 文件！", L"提示", MB_ICONWARNING);
        return;
    }

    try {
        bool addWm = (SendMessage(g_hChkAddWm, BM_GETCHECK, 0, 0) == BST_CHECKED);
        bool lockStruct = (SendMessage(g_hChkLockStruct, BM_GETCHECK, 0, 0) == BST_CHECKED);
        bool protectSheet = (SendMessage(g_hChkProtectSheet, BM_GETCHECK, 0, 0) == BST_CHECKED);
        bool unprotectAll = (SendMessage(g_hChkUnprotectAll, BM_GETCHECK, 0, 0) == BST_CHECKED);

        if (!addWm && !lockStruct && !protectSheet && !unprotectAll) {
            MessageBoxW(g_hMainWnd, L"请至少勾选一项操作（如：添加暗水印、锁定工作簿、工作表保护或解除保护）！", L"提示", MB_ICONWARNING);
            return;
        }

        wchar_t szWmText[256] = { 0 };
        wchar_t szTraceId[128] = { 0 };
        wchar_t szPassword[128] = { 0 };
        GetWindowTextW(g_hEditWmText, szWmText, 255);
        GetWindowTextW(g_hEditTraceId, szTraceId, 127);
        GetWindowTextW(g_hEditPassword, szPassword, 127);

        ExcelToolkitCore::ProcessOptions options;
        options.addWatermark = addWm;
        options.watermarkText = WStringToUtf8(szWmText);
        options.traceId = WStringToUtf8(szTraceId);
        options.lockWorkbookStructure = lockStruct;
        options.protectSheetContent = protectSheet;
        options.password = WStringToUtf8(szPassword);
        options.unprotectAll = unprotectAll;

        std::string inputPath = WStringToUtf8(g_selectedFilePath);
        std::string suffix = "_processed.xlsx";
        if (addWm && lockStruct) suffix = "_watermarked_protected.xlsx";
        else if (addWm) suffix = "_watermarked.xlsx";
        else if (unprotectAll) suffix = "_unprotected.xlsx";
        else if (lockStruct || protectSheet) suffix = "_protected.xlsx";

        size_t dotPos = inputPath.find_last_of('.');
        std::string outputPath = (dotPos != std::string::npos) ? (inputPath.substr(0, dotPos) + suffix) : (inputPath + suffix);

        SetWindowTextW(g_hStaticStatus, L"正在处理中，请稍候...");

        std::string err;
        bool ok = ExcelToolkitCore::ProcessExcelFile(inputPath, outputPath, options, err);

        if (ok) {
            std::wstring successMsg = L"处理完成！\n已生成新文件：\n" + Utf8ToWString(outputPath) + L"\n\n已应用项目：\n";
            if (addWm) successMsg += L"  - 添加 veryHidden 隐性暗水印\n";
            if (lockStruct) successMsg += L"  - 锁定工作簿结构（防止右键取消隐藏）\n";
            if (protectSheet) successMsg += L"  - 锁定工作表内容（防止单元格编辑）\n";
            if (unprotectAll) successMsg += L"  - 一键解除工作簿与工作表全部保护\n";

            SetWindowTextW(g_hStaticStatus, L"处理成功！输出文件已生成。");
            MessageBoxW(g_hMainWnd, successMsg.c_str(), L"操作成功", MB_ICONINFORMATION);
        } else {
            std::wstring errMsg = L"处理失败：\n" + Utf8ToWString(err);
            SetWindowTextW(g_hStaticStatus, L"处理失败，详情见弹窗。");
            MessageBoxW(g_hMainWnd, errMsg.c_str(), L"错误提示", MB_ICONERROR);
        }
    } catch (const std::exception& e) {
        std::wstring msg = L"处理发生异常: " + Utf8ToWString(e.what());
        MessageBoxW(g_hMainWnd, msg.c_str(), L"异常告警", MB_ICONERROR);
    } catch (...) {
        MessageBoxW(g_hMainWnd, L"发生未知内部异常，操作已安全终止。", L"异常告警", MB_ICONERROR);
    }
}

void OnExecuteInspect() {
    if (g_selectedFilePath.empty()) {
        MessageBoxW(g_hMainWnd, L"请先选择或拖入待取证的 Excel (.xlsx / .xls) 文件！", L"提示", MB_ICONWARNING);
        return;
    }

    try {
        std::string inputPath = WStringToUtf8(g_selectedFilePath);
        SetWindowTextW(g_hStaticStatus, L"正在解析底层工作簿 XML 结构...");

        ExcelToolkitCore::InspectWatermarkInfo info;
        std::string err;
        bool ok = ExcelToolkitCore::InspectWatermarkFile(inputPath, info, err);

        SendMessage(g_hListSheets, LB_RESETCONTENT, 0, 0);

        if (!ok) {
            std::wstring errMsg = L"解析失败: " + Utf8ToWString(err);
            SetWindowTextW(g_hStaticStatus, errMsg.c_str());
            SetWindowTextW(g_hEditInspectResult, errMsg.c_str());
            return;
        }

        for (const auto& s : info.sheets) {
            std::wstring line = L"[" + Utf8ToWString(s.state) + L"]  " + Utf8ToWString(s.name);
            if (s.state == "veryHidden") {
                line += L"  ◄ 【暗水印 carrier 深度隐藏】";
            }
            SendMessage(g_hListSheets, LB_ADDSTRING, 0, (LPARAM)line.c_str());
        }

        std::wstringstream wss;
        if (info.hasHiddenSheet && !info.traceId.empty()) {
            wss << L"【取证结果：检测到 veryHidden 隐性暗水印】\r\n"
                << L"溯源编号: " << Utf8ToWString(info.traceId) << L"\r\n"
                << L"水印文字: " << Utf8ToWString(info.watermarkText) << L"\r\n"
                << L"写入时间: " << Utf8ToWString(info.createdAt) << L"\r\n"
                << L"工作簿锁定: " << (info.isWorkbookLocked ? L"已锁定 (lockStructure=1)" : L"未锁定") << L"\r\n";
            SetWindowTextW(g_hStaticStatus, L"取证完成：成功检测并解码 veryHidden 暗水印！");
        } else {
            wss << L"【扫描结果：未发现 veryHidden 暗水印】\r\n"
                << L"工作簿工作表数量: " << info.sheets.size() << L"\r\n"
                << L"工作簿结构保护: " << (info.isWorkbookLocked ? L"已锁定" : L"未锁定") << L"\r\n";
            SetWindowTextW(g_hStaticStatus, L"扫描完成：未发现 veryHidden 隐藏工作表。");
        }

        SetWindowTextW(g_hEditInspectResult, wss.str().c_str());
    } catch (const std::exception& e) {
        std::wstring msg = L"取证发生异常: " + Utf8ToWString(e.what());
        MessageBoxW(g_hMainWnd, msg.c_str(), L"异常告警", MB_ICONERROR);
    } catch (...) {
        MessageBoxW(g_hMainWnd, L"取证过程发生未知异常。", L"异常告警", MB_ICONERROR);
    }
}

void UpdateTabVisibility() {
    g_currentTab = TabCtrl_GetCurSel(g_hTabCtrl);

    int showTab1 = (g_currentTab == 0) ? SW_SHOW : SW_HIDE;
    int showTab2 = (g_currentTab == 1) ? SW_SHOW : SW_HIDE;

    ShowWindow(g_hChkAddWm, showTab1);
    ShowWindow(g_hLblWmText, showTab1);
    ShowWindow(g_hEditWmText, showTab1);
    ShowWindow(g_hLblTraceId, showTab1);
    ShowWindow(g_hEditTraceId, showTab1);
    ShowWindow(g_hBtnGenTrace, showTab1);
    ShowWindow(g_hChkLockStruct, showTab1);
    ShowWindow(g_hChkProtectSheet, showTab1);
    ShowWindow(g_hLblPassword, showTab1);
    ShowWindow(g_hEditPassword, showTab1);
    ShowWindow(g_hChkUnprotectAll, showTab1);
    ShowWindow(g_hBtnExecuteProcess, showTab1);

    ShowWindow(g_hBtnStartInspect, showTab2);
    ShowWindow(g_hLblSheetsList, showTab2);
    ShowWindow(g_hListSheets, showTab2);
    ShowWindow(g_hLblInspectRes, showTab2);
    ShowWindow(g_hEditInspectResult, showTab2);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE: {
        g_hFont = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                              CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");

        g_hTabCtrl = CreateWindowExW(0, WC_TABCONTROLW, L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
                                     12, 12, 630, 430, hWnd, (HMENU)IDC_TABCONTROL, g_hInstance, NULL);
        SendMessage(g_hTabCtrl, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        TCITEMW tie = { 0 };
        tie.mask = TCIF_TEXT;
        tie.pszText = (LPWSTR)L"1. 综合安全处理 (暗水印 / 保护 / 解锁)";
        TabCtrl_InsertItem(g_hTabCtrl, 0, &tie);
        tie.pszText = (LPWSTR)L"2. 暗水印溯源与查看 (取证)";
        TabCtrl_InsertItem(g_hTabCtrl, 1, &tie);

        HWND hLblFile = CreateWindowExW(0, L"STATIC", L"目标 Excel 文件 (支持直接拖拽文件到此窗口):", WS_CHILD | WS_VISIBLE,
                                        28, 48, 380, 20, hWnd, NULL, g_hInstance, NULL);
        SendMessage(hLblFile, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hEditFilePath = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                          28, 70, 490, 26, hWnd, (HMENU)IDC_EDIT_FILEPATH, g_hInstance, NULL);
        SendMessage(g_hEditFilePath, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        HWND hBtnBrowse = CreateWindowExW(0, L"BUTTON", L"浏览...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                          526, 69, 90, 28, hWnd, (HMENU)IDC_BTN_BROWSE, g_hInstance, NULL);
        SendMessage(hBtnBrowse, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hChkAddWm = CreateWindowExW(0, L"BUTTON", L"[可选项] 写入 veryHidden 隐性暗水印工作表",
                                      WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                      28, 108, 380, 22, hWnd, (HMENU)IDC_CHK_ADD_WM, g_hInstance, NULL);
        SendMessage(g_hChkAddWm, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        SendMessage(g_hChkAddWm, BM_SETCHECK, BST_CHECKED, 0);

        g_hLblWmText = CreateWindowExW(0, L"STATIC", L"水印文字/说明:", WS_CHILD | WS_VISIBLE,
                                      48, 134, 110, 20, hWnd, NULL, g_hInstance, NULL);
        SendMessage(g_hLblWmText, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hEditWmText = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"绝密外发 · 专供交通运输局审计",
                                        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                        160, 132, 456, 24, hWnd, (HMENU)IDC_EDIT_WM_TEXT, g_hInstance, NULL);
        SendMessage(g_hEditWmText, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hLblTraceId = CreateWindowExW(0, L"STATIC", L"防伪溯源编号:", WS_CHILD | WS_VISIBLE,
                                       48, 164, 110, 20, hWnd, NULL, g_hInstance, NULL);
        SendMessage(g_hLblTraceId, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hEditTraceId = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"SEC-2026-WM888",
                                         WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                         160, 162, 350, 24, hWnd, (HMENU)IDC_EDIT_TRACE_ID, g_hInstance, NULL);
        SendMessage(g_hEditTraceId, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hBtnGenTrace = CreateWindowExW(0, L"BUTTON", L"随机生成", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                         518, 161, 98, 26, hWnd, (HMENU)IDC_BTN_GEN_TRACE, g_hInstance, NULL);
        SendMessage(g_hBtnGenTrace, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hChkLockStruct = CreateWindowExW(0, L"BUTTON", L"[可选项] 锁定工作簿结构 (lockStructure=1，禁止右键取消隐藏暗水印表)",
                                           WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                           28, 200, 580, 22, hWnd, (HMENU)IDC_CHK_LOCK_STRUCT, g_hInstance, NULL);
        SendMessage(g_hChkLockStruct, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        SendMessage(g_hChkLockStruct, BM_SETCHECK, BST_CHECKED, 0);

        g_hChkProtectSheet = CreateWindowExW(0, L"BUTTON", L"[可选项] 锁定工作表内容 (启用单元格只读防篡改保护)",
                                            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                            28, 228, 380, 22, hWnd, (HMENU)IDC_CHK_PROTECT_SHEET, g_hInstance, NULL);
        SendMessage(g_hChkProtectSheet, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hLblPassword = CreateWindowExW(0, L"STATIC", L"保护密码(可选):", WS_CHILD | WS_VISIBLE,
                                        410, 230, 100, 20, hWnd, NULL, g_hInstance, NULL);
        SendMessage(g_hLblPassword, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hEditPassword = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_PASSWORD | ES_AUTOHSCROLL,
                                          510, 228, 106, 24, hWnd, (HMENU)IDC_EDIT_PASSWORD, g_hInstance, NULL);
        SendMessage(g_hEditPassword, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hChkUnprotectAll = CreateWindowExW(0, L"BUTTON", L"[解锁模式] 一键解除工作簿与工作表全部保护限制",
                                            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                            28, 260, 580, 22, hWnd, (HMENU)IDC_CHK_UNPROTECT_ALL, g_hInstance, NULL);
        SendMessage(g_hChkUnprotectAll, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hBtnExecuteProcess = CreateWindowExW(0, L"BUTTON", L"执行所选操作 (生成处理后 Excel 文件)",
                                             WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                                             28, 300, 588, 38, hWnd, (HMENU)IDC_BTN_EXECUTE_PROCESS, g_hInstance, NULL);
        SendMessage(g_hBtnExecuteProcess, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hBtnStartInspect = CreateWindowExW(0, L"BUTTON", L"开始深度解析与取证", WS_CHILD | BS_PUSHBUTTON,
                                            28, 108, 588, 30, hWnd, (HMENU)IDC_BTN_START_INSPECT, g_hInstance, NULL);
        SendMessage(g_hBtnStartInspect, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hLblSheetsList = CreateWindowExW(0, L"STATIC", L"工作簿所有工作表及可见性状态:", WS_CHILD,
                                          28, 146, 300, 20, hWnd, NULL, g_hInstance, NULL);
        SendMessage(g_hLblSheetsList, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hListSheets = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"", WS_CHILD | WS_VSCROLL | LBS_NOTIFY,
                                        28, 168, 588, 90, hWnd, (HMENU)IDC_LIST_SHEETS, g_hInstance, NULL);
        SendMessage(g_hListSheets, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hLblInspectRes = CreateWindowExW(0, L"STATIC", L"暗水印特征解码与取证详情:", WS_CHILD,
                                          28, 266, 300, 20, hWnd, NULL, g_hInstance, NULL);
        SendMessage(g_hLblInspectRes, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hEditInspectResult = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"尚未扫描，请拖入文件或点击上方按钮开始解析。",
                                               WS_CHILD | ES_MULTILINE | ES_READONLY | WS_VSCROLL,
                                               28, 288, 588, 70, hWnd, (HMENU)IDC_EDIT_INSPECT_RESULT, g_hInstance, NULL);
        SendMessage(g_hEditInspectResult, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        g_hStaticStatus = CreateWindowExW(WS_EX_STATICEDGE, L"STATIC", L"就绪。请选择或拖入 Excel 文件。", WS_CHILD | WS_VISIBLE,
                                          12, 450, 630, 26, hWnd, (HMENU)IDC_STATIC_STATUS, g_hInstance, NULL);
        SendMessage(g_hStaticStatus, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        DragAcceptFiles(hWnd, TRUE);
        UpdateTabVisibility();
        break;
    }

    case WM_DROPFILES: {
        HDROP hDrop = (HDROP)wParam;
        wchar_t szFile[MAX_PATH] = { 0 };
        if (DragQueryFileW(hDrop, 0, szFile, MAX_PATH)) {
            g_selectedFilePath = szFile;
            SetWindowTextW(g_hEditFilePath, szFile);
            SetWindowTextW(g_hStaticStatus, (std::wstring(L"已拖入文件: ") + szFile).c_str());
            if (g_currentTab == 1) {
                OnExecuteInspect();
            }
        }
        DragFinish(hDrop);
        break;
    }

    case WM_NOTIFY: {
        LPNMHDR pnmhdr = (LPNMHDR)lParam;
        if (pnmhdr->idFrom == IDC_TABCONTROL && pnmhdr->code == TCN_SELCHANGE) {
            UpdateTabVisibility();
        }
        break;
    }

    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case IDC_BTN_BROWSE:
            OnBrowseFile();
            break;
        case IDC_BTN_GEN_TRACE:
            SetWindowTextW(g_hEditTraceId, GenerateTraceId().c_str());
            break;
        case IDC_CHK_UNPROTECT_ALL: {
            bool unprotectChecked = (SendMessage(g_hChkUnprotectAll, BM_GETCHECK, 0, 0) == BST_CHECKED);
            if (unprotectChecked) {
                SendMessage(g_hChkAddWm, BM_SETCHECK, BST_UNCHECKED, 0);
                SendMessage(g_hChkLockStruct, BM_SETCHECK, BST_UNCHECKED, 0);
                SendMessage(g_hChkProtectSheet, BM_SETCHECK, BST_UNCHECKED, 0);
            }
            break;
        }
        case IDC_CHK_ADD_WM:
        case IDC_CHK_LOCK_STRUCT:
        case IDC_CHK_PROTECT_SHEET: {
            if (SendMessage((HWND)lParam, BM_GETCHECK, 0, 0) == BST_CHECKED) {
                SendMessage(g_hChkUnprotectAll, BM_SETCHECK, BST_UNCHECKED, 0);
            }
            break;
        }
        case IDC_BTN_EXECUTE_PROCESS:
            OnExecuteProcess();
            break;
        case IDC_BTN_START_INSPECT:
            OnExecuteInspect();
            break;
        }
        break;
    }

    case WM_DESTROY:
        if (g_hFont) DeleteObject(g_hFont);
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    g_hInstance = hInstance;

    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_TAB_CLASSES | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icex);

    WNDCLASSEXW wcex = { 0 };
    wcex.cbSize = sizeof(WNDCLASSEXW);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(1));
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wcex.lpszClassName = L"ExcelToolkitGUI";
    RegisterClassExW(&wcex);

    g_hMainWnd = CreateWindowW(L"ExcelToolkitGUI", L"excel-toolkit — Excel 暗水印与保护工具 (原生独立 GUI 版)",
                               WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                               CW_USEDEFAULT, CW_USEDEFAULT, 672, 530, NULL, NULL, hInstance, NULL);

    if (!g_hMainWnd) return 0;

    ShowWindow(g_hMainWnd, nCmdShow);
    UpdateWindow(g_hMainWnd);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return (int)msg.wParam;
}
