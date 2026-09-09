/**
 * ============================================================================
 * excel-toolkit (gui_main.cpp) — Excel工具箱 (Windows 10/11 现代原生 Fluent 独立版)
 * 
 * 核心特性：
 *   - 现代 Windows 10/11 Fluent Design 视觉风格 (纯净白/浅灰分层底色，告别 Win98 土灰框)
 *   - 现代扁平化选项卡切换 (Segmented Tabs)、Excel 品牌翡翠绿主操作按钮 (自绘圆角扁平)
 *   - 全面支持 Per-Monitor V2 矢量级高 DPI 动态自适应缩放 (100%, 125%, 150%, 175%, 200%+)
 *   - 完美处理 WM_CTLCOLOR 消息，彻底消除控件灰底补丁与 3D 凹凸阴影
 *   - 纯原生 Win32 + Common Controls v6 界面，无任何控制台黑框 (/SUBSYSTEM:WINDOWS)
 *   - 卡片 1【加工】：写入 veryHidden 隐性暗水印、工作簿结构加锁、工作表只读保护
 *   - 卡片 2【脱壳】：深度解析工作表状态、提取 veryHidden 暗水印与溯源编号、一键脱壳解锁
 *   - 支持全窗口直接拖拽 Excel (.xlsx / .xls) 文件
 * ============================================================================
 */

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <uxtheme.h>
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
#pragma comment(lib, "uxtheme.lib")

// 控件 ID
#define IDC_TAB_BTN_1            1001
#define IDC_TAB_BTN_2            1002
#define IDC_EDIT_FILEPATH        1003
#define IDC_BTN_BROWSE           1004

// Tab 1 (加工)
#define IDC_CHK_ADD_WM           1005
#define IDC_EDIT_WM_TEXT         1006
#define IDC_EDIT_TRACE_ID        1007
#define IDC_BTN_GEN_TRACE        1008
#define IDC_CHK_LOCK_STRUCT      1009
#define IDC_CHK_PROTECT_SHEET    1010
#define IDC_EDIT_PASSWORD        1011
#define IDC_BTN_EXECUTE_PROCESS  1012

// Tab 2 (脱壳)
#define IDC_BTN_START_INSPECT    1020
#define IDC_LIST_SHEETS          1021
#define IDC_EDIT_INSPECT_RESULT  1022
#define IDC_BTN_QUICK_UNPROTECT  1023

// 状态条
#define IDC_STATIC_STATUS        1030

#define APP_VERSION_STR          L"v0.1.1"
#define APP_TITLE_STR            L"Excel 工具箱 v0.1.1"

// 现代配色规范 (Windows 10 / Fluent UI Palette)
#define COLOR_WIN_BG             RGB(243, 244, 246)  // 现代浅灰底色 #F3F4F6
#define COLOR_CARD_BG            RGB(255, 255, 255)  // 纯白卡片底色 #FFFFFF
#define COLOR_CARD_BORDER        RGB(229, 231, 235)  // 极细浅灰边框 #E5E7EB
#define COLOR_TEXT_MAIN          RGB(31, 41, 55)     // 主文字深灰 #1F2937
#define COLOR_TEXT_MUTED         RGB(107, 114, 128)  // 辅助说明灰 #6B7280
#define COLOR_EXCEL_GREEN        RGB(16, 124, 65)    // Excel 翡翠绿 #107C41
#define COLOR_EXCEL_GREEN_HOVER  RGB(14, 107, 56)    // 悬停深绿 #0E6B38
#define COLOR_EXCEL_GREEN_PRESS  RGB(11, 85, 44)     // 按下暗绿 #0B552C
#define COLOR_BLUE_ACCENT        RGB(0, 120, 212)    // 现代深青蓝 #0078D4
#define COLOR_BLUE_HOVER         RGB(16, 110, 190)   // 悬停蓝 #106EBE
#define COLOR_BLUE_PRESS         RGB(0, 90, 158)     // 按下蓝 #005A9E
#define COLOR_EDIT_BG            RGB(248, 250, 252)  // 输入框精致浅灰底纹 #F8FAFC
#define COLOR_EDIT_TEXT          RGB(15, 23, 42)     // 输入框清晰深色文字 #0F172A
#define COLOR_EDIT_BORDER        RGB(209, 213, 219)  // 输入框细边框 #D1D5DB

// DPI 缩放引擎
static int g_dpi = 96;

inline int Scale(int val) {
    return MulDiv(val, g_dpi, 96);
}

// 全局句柄与画刷
HINSTANCE g_hInstance = NULL;
HWND g_hMainWnd = NULL;
HBRUSH g_hBrushWinBg = NULL;
HBRUSH g_hBrushCardBg = NULL;
HBRUSH g_hBrushEditBg = NULL;

// 顶部现代化 Tab 按钮
HWND g_hBtnTab1 = NULL;
HWND g_hBtnTab2 = NULL;

// 文件选择区
HWND g_hLblFile = NULL;
HWND g_hEditFilePath = NULL;
HWND g_hBtnBrowse = NULL;

// Tab 1【加工】控件 (水印默认：绝密，防伪溯源码后台自动基于选文件时间戳生成，无需界面占位)
HWND g_hChkAddWm = NULL;
HWND g_hLblWmText = NULL;
HWND g_hEditWmText = NULL;

HWND g_hChkLockStruct = NULL;
HWND g_hChkProtectSheet = NULL;
HWND g_hLblPassword = NULL;
HWND g_hEditPassword = NULL;
HWND g_hBtnExecuteProcess = NULL;

// 输入框外框矩形 (用于优雅绘制底纹与 1px 细边框，使内部文字绝对竖直居中)
RECT g_rcBoxFile = { 0 };
RECT g_rcBoxWm = { 0 };
RECT g_rcBoxPwd = { 0 };

// 每次选择文件后更新的防伪溯源码 (保持 SEC-{UNIX时间戳} 格式)
static std::wstring g_currentTraceId = L"";

std::wstring GenerateTraceId() {
    auto now = std::chrono::system_clock::now();
    auto sec = std::chrono::duration_cast<std::chrono::seconds>(now.time_since_epoch()).count();
    std::wstringstream wss;
    wss << L"SEC-" << sec;
    return wss.str();
}

// Tab 2【脱壳】控件 (已移除工作表状态清单列表框，全宽呈现溯源码与取证报告)
HWND g_hBtnStartInspect = NULL;
HWND g_hLblInspectRes = NULL;
HWND g_hEditInspectResult = NULL;
HWND g_hBtnQuickUnprotect = NULL;

// 底部状态提示
HWND g_hStaticStatus = NULL;

// 字体对象
HFONT g_hFontNormal = NULL;
HFONT g_hFontBold = NULL;
HFONT g_hFontTitle = NULL;
HFONT g_hFontHeaderTitle = NULL;

int g_currentTab = 0; // 0 = 加工, 1 = 脱壳
std::wstring g_selectedFilePath = L"";

// 字符串转换工具函数
std::wstring Utf8ToWString(const std::string& str) {
    if (str.empty()) return L"";
    int size = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);
    if (size <= 1) {
        int ansiSize = MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, NULL, 0);
        if (ansiSize <= 1) return L"";
        std::wstring result(ansiSize - 1, 0);
        MultiByteToWideChar(CP_ACP, 0, str.c_str(), -1, &result[0], ansiSize);
        return result;
    }
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

void UpdateDpi(HWND hWnd) {
    typedef UINT(WINAPI* GetDpiForWindowProc)(HWND);
    static GetDpiForWindowProc pGetDpiForWindow = (GetDpiForWindowProc)GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetDpiForWindow");
    
    if (pGetDpiForWindow && hWnd) {
        g_dpi = (int)pGetDpiForWindow(hWnd);
    } else {
        HDC hdc = GetDC(hWnd);
        if (hdc) {
            g_dpi = GetDeviceCaps(hdc, LOGPIXELSY);
            ReleaseDC(hWnd, hdc);
        } else {
            g_dpi = 96;
        }
    }
    if (g_dpi < 96) g_dpi = 96;
}

void RecreateFonts() {
    if (g_hFontNormal) DeleteObject(g_hFontNormal);
    if (g_hFontBold) DeleteObject(g_hFontBold);
    if (g_hFontTitle) DeleteObject(g_hFontTitle);
    if (g_hFontHeaderTitle) DeleteObject(g_hFontHeaderTitle);

    g_hFontNormal = CreateFontW(Scale(-13), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");
    g_hFontBold   = CreateFontW(Scale(-13), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");
    g_hFontTitle  = CreateFontW(Scale(-14), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                                DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");
    g_hFontHeaderTitle = CreateFontW(Scale(-17), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
                                     DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
                                     CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Microsoft YaHei UI");
}

void ApplyFontsToControls() {
    if (g_hBtnTab1) SendMessageW(g_hBtnTab1, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);
    if (g_hBtnTab2) SendMessageW(g_hBtnTab2, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);

    if (g_hLblFile) SendMessageW(g_hLblFile, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);
    if (g_hEditFilePath) SendMessageW(g_hEditFilePath, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
    if (g_hBtnBrowse) SendMessageW(g_hBtnBrowse, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

    if (g_hChkAddWm) SendMessageW(g_hChkAddWm, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);
    if (g_hLblWmText) SendMessageW(g_hLblWmText, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
    if (g_hEditWmText) SendMessageW(g_hEditWmText, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);

    if (g_hChkLockStruct) SendMessageW(g_hChkLockStruct, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
    if (g_hChkProtectSheet) SendMessageW(g_hChkProtectSheet, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
    if (g_hLblPassword) SendMessageW(g_hLblPassword, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
    if (g_hEditPassword) SendMessageW(g_hEditPassword, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
    if (g_hBtnExecuteProcess) SendMessageW(g_hBtnExecuteProcess, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);

    if (g_hBtnStartInspect) SendMessageW(g_hBtnStartInspect, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);
    if (g_hLblInspectRes) SendMessageW(g_hLblInspectRes, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);
    if (g_hEditInspectResult) SendMessageW(g_hEditInspectResult, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
    if (g_hBtnQuickUnprotect) SendMessageW(g_hBtnQuickUnprotect, WM_SETFONT, (WPARAM)g_hFontTitle, TRUE);

    if (g_hStaticStatus) SendMessageW(g_hStaticStatus, WM_SETFONT, (WPARAM)g_hFontNormal, TRUE);
}

// 现代扁平化布局排版
void RealignControls(int clientWidth, int clientHeight) {
    if (!g_hMainWnd || clientWidth <= 0 || clientHeight <= 0) return;

    int padX = Scale(22);
    int padY = Scale(18);
    int contentW = clientWidth - padX * 2;

    // 1. 顶部 Tab 导航栏
    int headerH = Scale(52);
    int tabBtnW = Scale(150);
    int tabBtnH = Scale(36);
    int tabY = headerH + Scale(8);

    SetWindowPos(g_hBtnTab1, NULL, padX, tabY, tabBtnW, tabBtnH, SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(g_hBtnTab2, NULL, padX + tabBtnW + Scale(8), tabY, tabBtnW, tabBtnH, SWP_NOZORDER | SWP_NOACTIVATE);

    // 2. 主卡片容器区域
    int cardY = tabY + tabBtnH + Scale(12);
    int cardH = clientHeight - cardY - Scale(46);
    int innerX = padX + Scale(20);
    int innerW = contentW - Scale(40);

    // 3. 目标文件选择行 (卡片内顶部)
    int fileLblY = cardY + Scale(16);
    SetWindowPos(g_hLblFile, NULL, innerX, fileLblY, innerW, Scale(20), SWP_NOZORDER | SWP_NOACTIVATE);

    int fileInputY = fileLblY + Scale(24);
    int browseBtnW = Scale(96);
    int inputH = Scale(32);
    int editFileW = innerW - browseBtnW - Scale(10);
    int editH = Scale(20);

    g_rcBoxFile.left = innerX;
    g_rcBoxFile.top = fileInputY;
    g_rcBoxFile.right = innerX + editFileW;
    g_rcBoxFile.bottom = fileInputY + inputH;

    int editFileY = fileInputY + (inputH - editH) / 2;
    SetWindowPos(g_hEditFilePath, NULL, innerX + Scale(8), editFileY, editFileW - Scale(16), editH, SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(g_hBtnBrowse, NULL, innerX + editFileW + Scale(10), fileInputY, browseBtnW, inputH, SWP_NOZORDER | SWP_NOACTIVATE);

    // 4. Tab 1【加工】控件布局 (像素级严格网格对齐，文字绝对竖直居中)
    int wmSecY = fileInputY + inputH + Scale(20);
    SetWindowPos(g_hChkAddWm, NULL, innerX, wmSecY, innerW, Scale(24), SWP_NOZORDER | SWP_NOACTIVATE);

    int lblW = Scale(108); // 统一标签宽度，确保所有冒号和输入框左边缘像素级严格对齐
    int inputX = innerX + lblW + Scale(8);
    int totalInputW = innerW - lblW - Scale(8);
    int rowInputH = Scale(32);

    // 水印内容行
    int rowWmY = wmSecY + Scale(32);
    g_rcBoxWm.left = inputX;
    g_rcBoxWm.top = rowWmY;
    g_rcBoxWm.right = inputX + totalInputW;
    g_rcBoxWm.bottom = rowWmY + rowInputH;

    int editWmY = rowWmY + (rowInputH - editH) / 2;
    SetWindowPos(g_hLblWmText, NULL, innerX, rowWmY + Scale(6), lblW, Scale(20), SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(g_hEditWmText, NULL, inputX + Scale(8), editWmY, totalInputW - Scale(16), editH, SWP_NOZORDER | SWP_NOACTIVATE);

    // 结构锁定与保护部分 (溯源码后台自动记录，无需界面展示)
    int protSecY = rowWmY + rowInputH + Scale(24);
    SetWindowPos(g_hChkLockStruct, NULL, innerX, protSecY, innerW, Scale(24), SWP_NOZORDER | SWP_NOACTIVATE);

    int protRow2Y = protSecY + Scale(28);
    SetWindowPos(g_hChkProtectSheet, NULL, innerX, protRow2Y, innerW, Scale(24), SWP_NOZORDER | SWP_NOACTIVATE);

    // 密码输入行 (标签充分预留宽度 Scale(130)，彻底解决“保护密码 (可选):”文字被遮挡的问题)
    int rowPwdY = protRow2Y + Scale(28);
    int pwdLblW = Scale(130);
    int pwdInputX = innerX + Scale(22) + pwdLblW + Scale(6);
    int pwdEditW = Scale(240);

    g_rcBoxPwd.left = pwdInputX;
    g_rcBoxPwd.top = rowPwdY;
    g_rcBoxPwd.right = pwdInputX + pwdEditW;
    g_rcBoxPwd.bottom = rowPwdY + rowInputH;

    int editPwdY = rowPwdY + (rowInputH - editH) / 2;
    SetWindowPos(g_hLblPassword, NULL, innerX + Scale(22), rowPwdY + Scale(6), pwdLblW, Scale(20), SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(g_hEditPassword, NULL, pwdInputX + Scale(8), editPwdY, pwdEditW - Scale(16), editH, SWP_NOZORDER | SWP_NOACTIVATE);

    // 主操作大按钮 1 (加工导出)
    int btnExecH = Scale(46);
    int btnExecY = cardY + cardH - btnExecH - Scale(16);
    SetWindowPos(g_hBtnExecuteProcess, NULL, innerX, btnExecY, innerW, btnExecH, SWP_NOZORDER | SWP_NOACTIVATE);

    // 5. Tab 2【脱壳】控件布局 (全宽平铺报告展示，无需工作表列表框，清爽高辨识度)
    int inspBtnY = fileInputY + inputH + Scale(14);
    int inspBtnH = Scale(38);
    SetWindowPos(g_hBtnStartInspect, NULL, innerX, inspBtnY, innerW, inspBtnH, SWP_NOZORDER | SWP_NOACTIVATE);

    int contentSecY = inspBtnY + inspBtnH + Scale(12);
    int contentSecH = btnExecY - contentSecY - Scale(14);

    // 暗水印取证与溯源码详情 (全宽占满 innerW，充足高度完整展示全部信息，无需滚动)
    SetWindowPos(g_hLblInspectRes, NULL, innerX, contentSecY, innerW, Scale(20), SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(g_hEditInspectResult, NULL, innerX, contentSecY + Scale(24), innerW, contentSecH - Scale(24), SWP_NOZORDER | SWP_NOACTIVATE);

    // 主操作大按钮 2 (一键脱壳)
    SetWindowPos(g_hBtnQuickUnprotect, NULL, innerX, btnExecY, innerW, btnExecH, SWP_NOZORDER | SWP_NOACTIVATE);

    // 6. 底部状态提示条
    int statusY = clientHeight - Scale(32);
    SetWindowPos(g_hStaticStatus, NULL, padX, statusY, contentW, Scale(24), SWP_NOZORDER | SWP_NOACTIVATE);

    InvalidateRect(g_hMainWnd, NULL, TRUE);
}

void UpdateTabVisibility() {
    int showTab1 = (g_currentTab == 0) ? SW_SHOW : SW_HIDE;
    int showTab2 = (g_currentTab == 1) ? SW_SHOW : SW_HIDE;

    // Tab 1
    ShowWindow(g_hChkAddWm, showTab1);
    ShowWindow(g_hLblWmText, showTab1);
    ShowWindow(g_hEditWmText, showTab1);
    ShowWindow(g_hChkLockStruct, showTab1);
    ShowWindow(g_hChkProtectSheet, showTab1);
    ShowWindow(g_hLblPassword, showTab1);
    ShowWindow(g_hEditPassword, showTab1);
    ShowWindow(g_hBtnExecuteProcess, showTab1);

    // Tab 2
    ShowWindow(g_hBtnStartInspect, showTab2);
    ShowWindow(g_hLblInspectRes, showTab2);
    ShowWindow(g_hEditInspectResult, showTab2);
    ShowWindow(g_hBtnQuickUnprotect, showTab2);

    InvalidateRect(g_hBtnTab1, NULL, TRUE);
    InvalidateRect(g_hBtnTab2, NULL, TRUE);
    InvalidateRect(g_hMainWnd, NULL, TRUE);
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
            g_currentTraceId = GenerateTraceId(); // 每次选择文件后更新防伪溯源码时间戳: SEC-{UNIX时间戳}
            SetWindowTextW(g_hEditFilePath, szFile);
            SetWindowTextW(g_hStaticStatus, L"● 就绪：已载入目标文件。");
            if (g_currentTab == 1) {
                SendMessageW(g_hMainWnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_START_INSPECT, BN_CLICKED), 0);
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
        bool addWm = (SendMessageW(g_hChkAddWm, BM_GETCHECK, 0, 0) == BST_CHECKED);
        bool lockStruct = (SendMessageW(g_hChkLockStruct, BM_GETCHECK, 0, 0) == BST_CHECKED);
        bool protectSheet = (SendMessageW(g_hChkProtectSheet, BM_GETCHECK, 0, 0) == BST_CHECKED);

        if (!addWm && !lockStruct && !protectSheet) {
            MessageBoxW(g_hMainWnd, L"请至少勾选一项加工操作（如：添加暗水印、锁定工作簿结构或保护工作表内容）！", L"提示", MB_ICONWARNING);
            return;
        }

        wchar_t szWmText[256] = { 0 };
        wchar_t szPassword[128] = { 0 };
        GetWindowTextW(g_hEditWmText, szWmText, 255);
        GetWindowTextW(g_hEditPassword, szPassword, 127);

        // 防伪溯源码后台保持为当前选择文件的时间戳 SEC-{UNIX时间戳}
        if (g_currentTraceId.empty()) {
            g_currentTraceId = GenerateTraceId();
        }

        ExcelToolkitCore::ProcessOptions options;
        options.addWatermark = addWm;
        options.watermarkText = WStringToUtf8(szWmText);
        options.traceId = WStringToUtf8(g_currentTraceId);
        options.lockWorkbookStructure = lockStruct;
        options.protectSheetContent = protectSheet;
        options.password = WStringToUtf8(szPassword);
        options.unprotectAll = false;

        std::string inputPath = WStringToUtf8(g_selectedFilePath);
        std::string suffix = "_processed.xlsx";
        if (addWm && lockStruct) suffix = "_watermarked_protected.xlsx";
        else if (addWm) suffix = "_watermarked.xlsx";
        else if (lockStruct || protectSheet) suffix = "_protected.xlsx";

        size_t dotPos = inputPath.find_last_of('.');
        std::string outputPath = (dotPos != std::string::npos) ? (inputPath.substr(0, dotPos) + suffix) : (inputPath + suffix);

        SetWindowTextW(g_hStaticStatus, L"● 正在加工处理中，请稍候...");

        std::string err;
        bool ok = ExcelToolkitCore::ProcessExcelFile(inputPath, outputPath, options, err);

        if (ok) {
            std::wstring successMsg = L"加工成功！\n已生成加密与保护文件：\n" + Utf8ToWString(outputPath) + L"\n\n已应用项目：\n";
            if (addWm) successMsg += L"  ✔ 写入 veryHidden 隐性暗水印 (防伪溯源码: " + g_currentTraceId + L")\n";
            if (lockStruct) successMsg += L"  ✔ 锁定工作簿结构（禁止取消隐藏）\n";
            if (protectSheet) successMsg += L"  ✔ 锁定工作表内容（启用只读保护）\n";

            SetWindowTextW(g_hStaticStatus, L"● 加工完成！已成功输出新文件。");
            MessageBoxW(g_hMainWnd, successMsg.c_str(), L"操作成功", MB_ICONINFORMATION);
        } else {
            std::wstring errMsg = L"加工失败：\n" + Utf8ToWString(err);
            SetWindowTextW(g_hStaticStatus, L"● 加工失败，详情见弹窗提示。");
            MessageBoxW(g_hMainWnd, errMsg.c_str(), L"错误提示", MB_ICONERROR);
        }
    } catch (const std::exception& e) {
        std::wstring msg = L"加工发生异常: " + Utf8ToWString(e.what());
        MessageBoxW(g_hMainWnd, msg.c_str(), L"异常告警", MB_ICONERROR);
    } catch (...) {
        MessageBoxW(g_hMainWnd, L"发生未知内部异常，操作已安全终止。", L"异常告警", MB_ICONERROR);
    }
}

void OnExecuteInspect() {
    if (g_selectedFilePath.empty()) {
        MessageBoxW(g_hMainWnd, L"请先选择或拖入待脱壳/取证的 Excel (.xlsx / .xls) 文件！", L"提示", MB_ICONWARNING);
        return;
    }

    try {
        std::string inputPath = WStringToUtf8(g_selectedFilePath);
        SetWindowTextW(g_hStaticStatus, L"● 正在深度扫描底层工作簿 XML 结构...");

        ExcelToolkitCore::InspectWatermarkInfo info;
        std::string err;
        bool ok = ExcelToolkitCore::InspectWatermarkFile(inputPath, info, err);

        if (!ok) {
            std::wstring errMsg = L"解析失败: " + Utf8ToWString(err);
            SetWindowTextW(g_hStaticStatus, errMsg.c_str());
            SetWindowTextW(g_hEditInspectResult, errMsg.c_str());
            return;
        }

        // 统计工作表信息
        int visibleCount = 0;
        int hiddenCount = 0;
        int veryHiddenCount = 0;
        for (const auto& s : info.sheets) {
            if (s.state == "veryHidden") veryHiddenCount++;
            else if (s.state == "hidden") hiddenCount++;
            else visibleCount++;
        }

        std::wstringstream wss;
        if (!info.traceId.empty()) {
            wss << L"【扫描状态：成功捕获 veryHidden 隐性暗水印】\r\n"
                << L"==================================================\r\n"
                << L"【防伪溯源码】: " << Utf8ToWString(info.traceId) << L"\r\n"
                << L"--------------------------------------------------\r\n"
                << L"暗水印文字   : " << (info.watermarkText.empty() ? L"绝密" : Utf8ToWString(info.watermarkText)) << L"\r\n"
                << L"打标时间戳   : " << (info.createdAt.empty() ? L"(未记录)" : Utf8ToWString(info.createdAt)) << L"\r\n"
                << L"工作簿结构锁 : " << (info.isWorkbookLocked ? L"已锁定结构 (lockStructure=1，禁止右键取消隐藏)" : L"未锁定") << L"\r\n"
                << L"工作表统计   : 共 " << info.sheets.size() << L" 个 ("
                << visibleCount << L" 可见, " << veryHiddenCount << L" 深度隐藏暗水印表)\r\n"
                << L"==================================================\r\n"
                << L"💡 提示：如需清除限制，可点击下方【移除所有保护】按钮。";
            SetWindowTextW(g_hStaticStatus, (std::wstring(L"● 读取成功！防伪溯源码: ") + Utf8ToWString(info.traceId)).c_str());
        } else if (info.hasHiddenSheet) {
            wss << L"【扫描状态：检测到 veryHidden 隐性暗水印工作表】\r\n"
                << L"==================================================\r\n"
                << L"【防伪溯源码】: (未提取到标准编号)\r\n"
                << L"--------------------------------------------------\r\n"
                << L"暗水印文字   : " << (info.watermarkText.empty() ? L"(未提取到)" : Utf8ToWString(info.watermarkText)) << L"\r\n"
                << L"工作簿结构锁 : " << (info.isWorkbookLocked ? L"已锁定结构 (lockStructure=1)" : L"未锁定") << L"\r\n"
                << L"工作表统计   : 共 " << info.sheets.size() << L" 个 ("
                << visibleCount << L" 可见, " << veryHiddenCount << L" 深度隐藏暗水印表)\r\n"
                << L"==================================================\r\n"
                << L"💡 提示：如需清除限制，可点击下方【移除所有保护】按钮。";
            SetWindowTextW(g_hStaticStatus, L"● 读取完成：检测到 veryHidden 深度隐藏工作表。");
        } else {
            wss << L"【扫描状态：未发现 veryHidden 暗水印】\r\n"
                << L"==================================================\r\n"
                << L"【防伪溯源码】: (该文件无防伪溯源码)\r\n"
                << L"--------------------------------------------------\r\n"
                << L"工作表总数   : " << info.sheets.size() << L" 个 (" << visibleCount << L" 个正常可见)\r\n"
                << L"工作簿保护   : " << (info.isWorkbookLocked ? L"已锁定 (需脱壳)" : L"未锁定") << L"\r\n"
                << L"==================================================\r\n"
                << L"💡 提示：该文件未嵌入隐性暗水印。若有只读密码保护，可直接点击下方【移除所有保护】按钮。";
            SetWindowTextW(g_hStaticStatus, L"● 扫描完成：工作簿未发现隐性暗水印。");
        }

        SetWindowTextW(g_hEditInspectResult, wss.str().c_str());
    } catch (const std::exception& e) {
        std::wstring msg = L"脱壳取证异常: " + Utf8ToWString(e.what());
        MessageBoxW(g_hMainWnd, msg.c_str(), L"异常告警", MB_ICONERROR);
    } catch (...) {
        MessageBoxW(g_hMainWnd, L"脱壳取证发生未知异常。", L"异常告警", MB_ICONERROR);
    }
}

void OnExecuteQuickUnprotect() {
    if (g_selectedFilePath.empty()) {
        MessageBoxW(g_hMainWnd, L"请先选择或拖入待脱壳解锁的 Excel (.xlsx / .xls) 文件！", L"提示", MB_ICONWARNING);
        return;
    }

    try {
        std::string inputPath = WStringToUtf8(g_selectedFilePath);
        std::string suffix = "_unlocked.xlsx";
        size_t dotPos = inputPath.find_last_of('.');
        std::string outputPath = (dotPos != std::string::npos) ? (inputPath.substr(0, dotPos) + suffix) : (inputPath + suffix);

        SetWindowTextW(g_hStaticStatus, L"● 正在一键脱壳并清除全部保护与密码限制...");

        ExcelToolkitCore::ProcessOptions options;
        options.unprotectAll = true;
        std::string err;

        bool ok = ExcelToolkitCore::ProcessExcelFile(inputPath, outputPath, options, err);

        if (ok) {
            std::wstring msg = L"一键脱壳解锁成功！\n已生成纯净未加密版 Excel 文件：\n" + Utf8ToWString(outputPath) + L"\n\n已清除项：\n  ✔ 移除工作簿结构锁定 (lockStructure)\n  ✔ 移除所有工作表保护与只读哈希密码 (sheetProtection)\n  ✔ 恢复所有工作表隐藏权限";
            SetWindowTextW(g_hStaticStatus, L"● 脱壳成功！已生成纯净解锁版文件。");
            MessageBoxW(g_hMainWnd, msg.c_str(), L"脱壳成功", MB_ICONINFORMATION);
            OnExecuteInspect();
        } else {
            std::wstring errMsg = L"脱壳失败：\n" + Utf8ToWString(err);
            SetWindowTextW(g_hStaticStatus, L"● 脱壳失败。");
            MessageBoxW(g_hMainWnd, errMsg.c_str(), L"错误提示", MB_ICONERROR);
        }
    } catch (const std::exception& e) {
        std::wstring msg = L"脱壳发生异常: " + Utf8ToWString(e.what());
        MessageBoxW(g_hMainWnd, msg.c_str(), L"异常告警", MB_ICONERROR);
    }
}

// 现代圆角扁平按钮自绘制
void DrawModernButton(LPDRAWITEMSTRUCT pDIS, COLORREF bgColor, COLORREF hoverColor, COLORREF pressColor, COLORREF textColor, const wchar_t* text) {
    HDC hdc = pDIS->hDC;
    RECT rc = pDIS->rcItem;
    bool isPressed = (pDIS->itemState & ODS_SELECTED);
    bool isDisabled = (pDIS->itemState & ODS_DISABLED);

    COLORREF currentBg = bgColor;
    if (isDisabled) {
        currentBg = RGB(220, 224, 230);
        textColor = RGB(156, 163, 175);
    } else if (isPressed) {
        currentBg = pressColor;
    }

    HBRUSH hBr = CreateSolidBrush(currentBg);
    HPEN hPen = CreatePen(PS_SOLID, 1, currentBg);
    HBRUSH hOldBr = (HBRUSH)SelectObject(hdc, hBr);
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

    int radius = Scale(6);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, radius, radius);

    SelectObject(hdc, hOldBr);
    SelectObject(hdc, hOldPen);
    DeleteObject(hBr);
    DeleteObject(hPen);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    HFONT hOldFont = (HFONT)SelectObject(hdc, g_hFontTitle);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, hOldFont);
}

// 现代 Segmented Tab 选项卡绘制
void DrawModernTabButton(LPDRAWITEMSTRUCT pDIS, int tabIndex, const wchar_t* text) {
    HDC hdc = pDIS->hDC;
    RECT rc = pDIS->rcItem;
    bool isSelected = (g_currentTab == tabIndex);

    COLORREF bg = isSelected ? COLOR_CARD_BG : RGB(235, 238, 242);
    COLORREF textClr = isSelected ? COLOR_EXCEL_GREEN : COLOR_TEXT_MUTED;

    HBRUSH hBr = CreateSolidBrush(bg);
    HPEN hPen = CreatePen(PS_SOLID, 1, isSelected ? COLOR_CARD_BORDER : RGB(225, 228, 232));
    HBRUSH hOldBr = (HBRUSH)SelectObject(hdc, hBr);
    HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);

    int radius = Scale(8);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, radius, radius);

    if (isSelected) {
        // 选中时底部高亮小指示条
        RECT rcIndicator = { rc.left + Scale(18), rc.bottom - Scale(3), rc.right - Scale(18), rc.bottom - Scale(1) };
        HBRUSH hIndBr = CreateSolidBrush(COLOR_EXCEL_GREEN);
        FillRect(hdc, &rcIndicator, hIndBr);
        DeleteObject(hIndBr);
    }

    SelectObject(hdc, hOldBr);
    SelectObject(hdc, hOldPen);
    DeleteObject(hBr);
    DeleteObject(hPen);

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textClr);
    HFONT hOldFont = (HFONT)SelectObject(hdc, isSelected ? g_hFontBold : g_hFontNormal);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, hOldFont);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_CREATE: {
        UpdateDpi(hWnd);
        RecreateFonts();

        g_hBrushWinBg = CreateSolidBrush(COLOR_WIN_BG);
        g_hBrushCardBg = CreateSolidBrush(COLOR_CARD_BG);
        g_hBrushEditBg = CreateSolidBrush(COLOR_EDIT_BG);

        // 1. 顶部 Segmented Tab 按钮 (自绘制)
        g_hBtnTab1 = CreateWindowExW(0, L"BUTTON", L"⚡ 加工与加锁",
                                     WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                                     0, 0, 100, 20, hWnd, (HMENU)IDC_TAB_BTN_1, g_hInstance, NULL);

        g_hBtnTab2 = CreateWindowExW(0, L"BUTTON", L"🔍 深度脱壳与取证",
                                     WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                                     0, 0, 100, 20, hWnd, (HMENU)IDC_TAB_BTN_2, g_hInstance, NULL);

        // 2. 公共：目标文件选择 (边框与底纹由外框精绘，文字绝对竖直居中)
        g_hLblFile = CreateWindowExW(0, L"STATIC", L"目标 Excel 文件 (支持直接拖拽文件到本窗口):", WS_CHILD | WS_VISIBLE,
                                     0, 0, 100, 20, hWnd, NULL, g_hInstance, NULL);

        g_hEditFilePath = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                          0, 0, 100, 20, hWnd, (HMENU)IDC_EDIT_FILEPATH, g_hInstance, NULL);

        g_hBtnBrowse = CreateWindowExW(0, L"BUTTON", L"浏览文件...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
                                       0, 0, 100, 20, hWnd, (HMENU)IDC_BTN_BROWSE, g_hInstance, NULL);

        // 3. Tab 1【加工】控件 (水印默认：绝密，溯源编号在打开瞬间固定且不可更改)
        g_hChkAddWm = CreateWindowExW(0, L"BUTTON", L"写入 veryHidden 隐性暗水印工作表",
                                      WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                      0, 0, 100, 20, hWnd, (HMENU)IDC_CHK_ADD_WM, g_hInstance, NULL);
        SendMessageW(g_hChkAddWm, BM_SETCHECK, BST_CHECKED, 0);

        g_hLblWmText = CreateWindowExW(0, L"STATIC", L"水印显示内容:", WS_CHILD | WS_VISIBLE,
                                       0, 0, 100, 20, hWnd, NULL, g_hInstance, NULL);

        g_hEditWmText = CreateWindowExW(0, L"EDIT", L"绝密",
                                         WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
                                         0, 0, 100, 20, hWnd, (HMENU)IDC_EDIT_WM_TEXT, g_hInstance, NULL);

        // 初始化默认溯源码 (格式 SEC-{UNIX时间戳})
        g_currentTraceId = GenerateTraceId();

        g_hChkLockStruct = CreateWindowExW(0, L"BUTTON", L"锁定工作簿结构 (禁止右键取消隐藏暗水印工作表)",
                                            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                            0, 0, 100, 20, hWnd, (HMENU)IDC_CHK_LOCK_STRUCT, g_hInstance, NULL);
        SendMessageW(g_hChkLockStruct, BM_SETCHECK, BST_CHECKED, 0);

        g_hChkProtectSheet = CreateWindowExW(0, L"BUTTON", L"锁定工作表内容 (启用单元格只读防篡改保护)",
                                             WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                                             0, 0, 100, 20, hWnd, (HMENU)IDC_CHK_PROTECT_SHEET, g_hInstance, NULL);

        g_hLblPassword = CreateWindowExW(0, L"STATIC", L"保护密码 (可选):", WS_CHILD | WS_VISIBLE,
                                         0, 0, 100, 20, hWnd, NULL, g_hInstance, NULL);

        g_hEditPassword = CreateWindowExW(0, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_PASSWORD | ES_AUTOHSCROLL,
                                           0, 0, 100, 20, hWnd, (HMENU)IDC_EDIT_PASSWORD, g_hInstance, NULL);

        g_hBtnExecuteProcess = CreateWindowExW(0, L"BUTTON", L"执行",
                                              WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                                              0, 0, 100, 20, hWnd, (HMENU)IDC_BTN_EXECUTE_PROCESS, g_hInstance, NULL);

        // 4. Tab 2【脱壳】控件
        g_hBtnStartInspect = CreateWindowExW(0, L"BUTTON", L"检测文件",
                                             WS_CHILD | BS_PUSHBUTTON,
                                             0, 0, 100, 20, hWnd, (HMENU)IDC_BTN_START_INSPECT, g_hInstance, NULL);

        g_hLblInspectRes = CreateWindowExW(0, L"STATIC", L"暗水印特征解码与安全取证详情:", WS_CHILD,
                                           0, 0, 100, 20, hWnd, NULL, g_hInstance, NULL);

        g_hEditInspectResult = CreateWindowExW(0, L"EDIT", L"尚未检测，请载入 Excel 文件并点击上方【检测文件】按钮开始解析。",
                                                WS_CHILD | WS_BORDER | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
                                                0, 0, 100, 20, hWnd, (HMENU)IDC_EDIT_INSPECT_RESULT, g_hInstance, NULL);

        g_hBtnQuickUnprotect = CreateWindowExW(0, L"BUTTON", L"移除所有保护",
                                               WS_CHILD | BS_OWNERDRAW,
                                               0, 0, 100, 20, hWnd, (HMENU)IDC_BTN_QUICK_UNPROTECT, g_hInstance, NULL);

        // 5. 底部状态栏
        g_hStaticStatus = CreateWindowExW(0, L"STATIC", L"● 就绪。请选择或直接拖入 Excel 文件到本窗口。", WS_CHILD | WS_VISIBLE,
                                           0, 0, 100, 20, hWnd, (HMENU)IDC_STATIC_STATUS, g_hInstance, NULL);

        ApplyFontsToControls();

        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        RealignControls(rcClient.right - rcClient.left, rcClient.bottom - rcClient.top);

        DragAcceptFiles(hWnd, TRUE);
        UpdateTabVisibility();
        break;
    }

    case WM_SIZE: {
        int width = LOWORD(lParam);
        int height = HIWORD(lParam);
        RealignControls(width, height);
        break;
    }

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);

        RECT rcClient;
        GetClientRect(hWnd, &rcClient);

        // 1. 窗口现代背景底色
        FillRect(hdc, &rcClient, g_hBrushWinBg);

        // 2. 现代顶部 Header Banner (纯白底色，带微灰分割线)
        int headerH = Scale(52);
        RECT rcHeader = { 0, 0, rcClient.right, headerH };
        HBRUSH hBrWhite = CreateSolidBrush(COLOR_CARD_BG);
        FillRect(hdc, &rcHeader, hBrWhite);

        // 绘制大标题 (现代 Fluent 极简左对齐排版，单行居中)
        int textX = Scale(22);

        SetBkMode(hdc, TRANSPARENT);
        HFONT hOldFont = (HFONT)SelectObject(hdc, g_hFontHeaderTitle);
        SetTextColor(hdc, COLOR_TEXT_MAIN);
        RECT rcTitle = { textX, 0, rcClient.right - Scale(20), headerH };
        DrawTextW(hdc, APP_TITLE_STR, -1, &rcTitle, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        // Header 下方细分割线
        HPEN hPenDivider = CreatePen(PS_SOLID, 1, COLOR_CARD_BORDER);
        HPEN hOldPen = (HPEN)SelectObject(hdc, hPenDivider);
        MoveToEx(hdc, 0, headerH, NULL);
        LineTo(hdc, rcClient.right, headerH);

        // 3. 中间内容卡片 (纯白现代圆角卡片)
        int padX = Scale(22);
        int cardY = headerH + Scale(56);
        int cardH = rcClient.bottom - cardY - Scale(46);
        RECT rcCard = { padX, cardY, rcClient.right - padX, cardY + cardH };

        HBRUSH hOldBr = (HBRUSH)SelectObject(hdc, hBrWhite);
        RoundRect(hdc, rcCard.left, rcCard.top, rcCard.right, rcCard.bottom, Scale(10), Scale(10));

        // 卡片内部细分线
        if (g_currentTab == 0) {
            int line1Y = cardY + Scale(68);
            MoveToEx(hdc, rcCard.left + Scale(20), line1Y, NULL);
            LineTo(hdc, rcCard.right - Scale(20), line1Y);

            int line2Y = cardY + Scale(154);
            MoveToEx(hdc, rcCard.left + Scale(20), line2Y, NULL);
            LineTo(hdc, rcCard.right - Scale(20), line2Y);
        }

        // 4. 绘制输入框的外框与浅灰底纹 (确保内嵌文本绝对竖直居中且边框平滑)
        auto DrawEditContainer = [&](const RECT& rc) {
            if (rc.right <= rc.left || rc.bottom <= rc.top) return;
            HBRUSH hBr = CreateSolidBrush(COLOR_EDIT_BG);
            HPEN hPen = CreatePen(PS_SOLID, 1, COLOR_EDIT_BORDER);
            HBRUSH hPrevBr = (HBRUSH)SelectObject(hdc, hBr);
            HPEN hPrevPen = (HPEN)SelectObject(hdc, hPen);
            int r = Scale(4);
            RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, r, r);
            SelectObject(hdc, hPrevBr);
            SelectObject(hdc, hPrevPen);
            DeleteObject(hBr);
            DeleteObject(hPen);
        };

        DrawEditContainer(g_rcBoxFile);
        if (g_currentTab == 0) {
            DrawEditContainer(g_rcBoxWm);
            DrawEditContainer(g_rcBoxPwd);
        }

        SelectObject(hdc, hOldBr);
        SelectObject(hdc, hOldPen);
        SelectObject(hdc, hOldFont);
        DeleteObject(hBrWhite);
        DeleteObject(hPenDivider);

        EndPaint(hWnd, &ps);
        break;
    }

    // 自绘制按钮 (现代圆角扁平主按钮与 Tab 导航)
    case WM_DRAWITEM: {
        LPDRAWITEMSTRUCT pDIS = (LPDRAWITEMSTRUCT)lParam;
        if (pDIS->CtlID == IDC_TAB_BTN_1) {
            DrawModernTabButton(pDIS, 0, L"⚡ 加工与加锁");
            return TRUE;
        } else if (pDIS->CtlID == IDC_TAB_BTN_2) {
            DrawModernTabButton(pDIS, 1, L"🔍 深度脱壳与取证");
            return TRUE;
        } else if (pDIS->CtlID == IDC_BTN_EXECUTE_PROCESS) {
            DrawModernButton(pDIS, COLOR_EXCEL_GREEN, COLOR_EXCEL_GREEN_HOVER, COLOR_EXCEL_GREEN_PRESS, RGB(255, 255, 255), L"执行");
            return TRUE;
        } else if (pDIS->CtlID == IDC_BTN_QUICK_UNPROTECT) {
            DrawModernButton(pDIS, COLOR_BLUE_ACCENT, COLOR_BLUE_HOVER, COLOR_BLUE_PRESS, RGB(255, 255, 255), L"移除所有保护");
            return TRUE;
        }
        break;
    }

    // 彻底消除控件灰底与凹凸 3D 边框 (现代白底无缝透明贴合)
    case WM_CTLCOLORSTATIC: {
        HDC hdcStatic = (HDC)wParam;
        HWND hCtl = (HWND)lParam;
        SetBkMode(hdcStatic, TRANSPARENT);
        
        if (hCtl == g_hStaticStatus) {
            SetTextColor(hdcStatic, COLOR_TEXT_MUTED);
            return (LRESULT)g_hBrushWinBg;
        }

        SetTextColor(hdcStatic, COLOR_TEXT_MAIN);
        return (LRESULT)g_hBrushCardBg;
    }

    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX: {
        HDC hdcCtl = (HDC)wParam;
        SetBkColor(hdcCtl, COLOR_EDIT_BG);
        SetTextColor(hdcCtl, COLOR_EDIT_TEXT);
        return (LRESULT)g_hBrushEditBg;
    }

    case WM_DPICHANGED: {
        g_dpi = HIWORD(wParam);
        RecreateFonts();
        ApplyFontsToControls();

        RECT* const prcNewWindow = (RECT*)lParam;
        SetWindowPos(hWnd, NULL,
                     prcNewWindow->left, prcNewWindow->top,
                     prcNewWindow->right - prcNewWindow->left,
                     prcNewWindow->bottom - prcNewWindow->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);

        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        RealignControls(rcClient.right - rcClient.left, rcClient.bottom - rcClient.top);
        break;
    }

    case WM_DROPFILES: {
        HDROP hDrop = (HDROP)wParam;
        wchar_t szFile[MAX_PATH] = { 0 };
        if (DragQueryFileW(hDrop, 0, szFile, MAX_PATH)) {
            g_selectedFilePath = szFile;
            g_currentTraceId = GenerateTraceId(); // 每次选择/拖入文件后更新防伪溯源码时间戳: SEC-{UNIX时间戳}
            SetWindowTextW(g_hEditFilePath, szFile);
            SetWindowTextW(g_hStaticStatus, (std::wstring(L"● 已载入拖入文件: ") + szFile).c_str());
            if (g_currentTab == 1) {
                OnExecuteInspect();
            }
        }
        DragFinish(hDrop);
        break;
    }

    case WM_LBUTTONDOWN: {
        POINT pt = { LOWORD(lParam), HIWORD(lParam) };
        if (PtInRect(&g_rcBoxFile, pt)) {
            SetFocus(g_hEditFilePath);
        } else if (g_currentTab == 0) {
            if (PtInRect(&g_rcBoxWm, pt)) SetFocus(g_hEditWmText);
            else if (PtInRect(&g_rcBoxPwd, pt)) SetFocus(g_hEditPassword);
        }
        break;
    }

    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case IDC_TAB_BTN_1:
            g_currentTab = 0;
            UpdateTabVisibility();
            break;
        case IDC_TAB_BTN_2:
            g_currentTab = 1;
            UpdateTabVisibility();
            if (!g_selectedFilePath.empty()) {
                OnExecuteInspect();
            }
            break;
        case IDC_BTN_BROWSE:
            OnBrowseFile();
            break;
        case IDC_BTN_EXECUTE_PROCESS:
            OnExecuteProcess();
            break;
        case IDC_BTN_START_INSPECT:
            OnExecuteInspect();
            break;
        case IDC_BTN_QUICK_UNPROTECT:
            OnExecuteQuickUnprotect();
            break;
        }
        break;
    }

    case WM_DESTROY:
        if (g_hBrushWinBg) DeleteObject(g_hBrushWinBg);
        if (g_hBrushCardBg) DeleteObject(g_hBrushCardBg);
        if (g_hBrushEditBg) DeleteObject(g_hBrushEditBg);
        if (g_hFontNormal) DeleteObject(g_hFontNormal);
        if (g_hFontBold) DeleteObject(g_hFontBold);
        if (g_hFontTitle) DeleteObject(g_hFontTitle);
        if (g_hFontHeaderTitle) DeleteObject(g_hFontHeaderTitle);
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProcW(hWnd, message, wParam, lParam);
    }
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow) {
    g_hInstance = hInstance;

    // 启用现代 Windows 10/11 动态 DPI 缩放 (Per-Monitor V2)
    typedef BOOL(WINAPI* SetProcessDpiAwarenessContextProc)(DPI_AWARENESS_CONTEXT);
    SetProcessDpiAwarenessContextProc pSetDpiContext = 
        (SetProcessDpiAwarenessContextProc)GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetProcessDpiAwarenessContext");
    
    if (pSetDpiContext) {
        pSetDpiContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    } else {
        SetProcessDPIAware();
    }

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
    wcex.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(1));
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = NULL; // 使用自定义双缓冲/WM_PAINT 纯净背景
    wcex.lpszClassName = L"ExcelToolkitGUI";
    RegisterClassExW(&wcex);

    HDC hdcScreen = GetDC(NULL);
    if (hdcScreen) {
        g_dpi = GetDeviceCaps(hdcScreen, LOGPIXELSY);
        ReleaseDC(NULL, hdcScreen);
    }
    if (g_dpi < 96) g_dpi = 96;

    int initialWidth  = Scale(780);
    int initialHeight = Scale(650);

    g_hMainWnd = CreateWindowW(L"ExcelToolkitGUI", APP_TITLE_STR,
                               WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
                               CW_USEDEFAULT, CW_USEDEFAULT, initialWidth, initialHeight, NULL, NULL, hInstance, NULL);

    if (!g_hMainWnd) return 0;

    ShowWindow(g_hMainWnd, nCmdShow);
    UpdateWindow(g_hMainWnd);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    return (int)msg.wParam;
}
