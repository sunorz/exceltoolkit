/**
 * ============================================================================
 * excel_toolkit_core.h - 核心处理接口 (暗水印、溯源查询、工作簿/工作表保护)
 * ============================================================================
 */

#pragma once

#include <string>
#include <vector>
#include <cstdint>

inline const std::string HIDDEN_SHEET_NAME_CONST = "~sysdata_cfg";

class ExcelToolkitCore {
public:
    struct SheetMeta {
        std::string name;
        std::string state; // "visible", "hidden", "veryHidden"
        int sheetId = 0;
    };

    struct InspectWatermarkInfo {
        std::vector<SheetMeta> sheets;
        bool hasHiddenSheet = false;
        std::string traceId;
        std::string watermarkText;
        std::string createdAt;
        bool isWorkbookLocked = false;
    };

    // 综合处理选项结构体（支持自由勾选 Checkbox 组合）
    struct ProcessOptions {
        bool addWatermark = true;                 // 是否添加暗水印 (veryHidden)
        std::string watermarkText = "";           // 暗水印文字
        std::string issuer = "ADMIN";             // 签发人（中文或英文）
        std::string traceId = "";                 // 溯源编号格式：{签发人}-{UNIX时间戳}
        std::string sheetName = "~sysdata_cfg";   // 隐性工作表名称
        
        bool lockWorkbookStructure = true;        // 是否锁定工作簿结构（防止取消隐藏）
        
        // 保护水印选项：专用于防止用户删除表格里的实体可见水印，同时完全放开表格内文字编辑
        bool addVisibleWatermark = true;
        
        bool protectSheetContent = true;
        std::string visibleWatermarkText = "";
        bool unlockAllCells = false;          // 保护水印开关
            // 实体可见水印（铺满全表含无内容区域，打印可见）
        std::string password = "";                // 保护密码（可选）
        
        bool unprotectAll = false;                // 是否一键解除所有保护（工作簿与工作表）
    };

    // 统一处理执行入口（根据 Checkbox 选项执行相应组合）
    static bool ProcessExcelFile(
        const std::string& inputPath,
        const std::string& outputPath,
        const ProcessOptions& options,
        std::string& outError
    );

    // 功能 1：添加暗水印 (兼容旧接口)
    static bool EmbedWatermark(
        const std::string& filePath,
        const std::string& watermarkText,
        const std::string& traceId,
        std::string& outError
    );

    // 功能 2：查看暗水印与结构全景
    static bool InspectWatermarkFile(
        const std::string& filePath,
        InspectWatermarkInfo& outInfo,
        std::string& outError
    );

    // 功能 3：开启保护
    static bool ProtectWorkbook(
        const std::string& filePath,
        const std::string& password,
        std::string& outError
    );

    // 功能 4：解除保护
    static bool UnprotectWorkbook(
        const std::string& filePath,
        std::string& outError
    );

    // 辅助工具
    static bool CheckSofficeAvailable();
    static std::string ConvertWithSoffice(const std::string& srcPath, const std::string& targetExt, const std::string& workDir);
    static std::string GenerateRandomHex8();
    static std::string GetIsoTimestampNow();
    static uint16_t ComputeExcelPasswordHash(const std::string& password);
};
