/**
 * ============================================================================
 * excel_toolkit_core.h - 核心处理接口 (暗水印、溯源查询、工作表/工作簿保护)
 * ============================================================================
 */

#pragma once

#include <string>
#include <vector>

inline const std::string HIDDEN_SHEET_NAME_CONST = "~sysdata_cfg";

class ExcelToolkitCore {
public:
    struct SheetMeta {
        std::string name;
        std::string state; // "visible", "hidden", "veryHidden"
    };

    struct InspectWatermarkInfo {
        std::vector<SheetMeta> sheets;
        bool hasHiddenSheet = false;
        std::string traceId;
        std::string watermarkText;
        std::string createdAt;
    };

    // 功能 1：添加暗水印
    static bool EmbedWatermark(
        const std::string& filePath,
        const std::string& watermarkText,
        const std::string& traceId,
        std::string& outError
    );

    // 功能 2：查看暗水印
    static bool InspectWatermarkFile(
        const std::string& filePath,
        InspectWatermarkInfo& outInfo,
        std::string& outError
    );

    // 功能 3：开启保护（锁定单元格编辑、锁定工作簿结构防止取消隐藏暗水印表）
    static bool ProtectWorkbook(
        const std::string& filePath,
        const std::string& password,
        std::string& outError
    );

    // 功能 3：解除保护
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
