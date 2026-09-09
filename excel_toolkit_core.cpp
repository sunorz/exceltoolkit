/**
 * ============================================================================
 * excel_toolkit_core.cpp - OpenXML / ZIP 深度实现
 * ============================================================================
 */

#define _CRT_SECURE_NO_WARNINGS
#include "excel_toolkit_core.h"
#include "miniz.h"

#include <iostream>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <chrono>
#include <iomanip>
#include <random>
#include <cstdlib>
#include <filesystem>
#include <map>
#include <set>

namespace fs = std::filesystem;

// 生成 8 位十六进制随机 ID (对应 uuid.uuid4().hex[:8])
std::string ExcelToolkitCore::GenerateRandomHex8() {
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<uint32_t> dis(0, 0xFFFFFFFF);
    uint32_t val = dis(gen);
    std::stringstream ss;
    ss << std::hex << std::setw(8) << std::setfill('0') << val;
    return ss.str();
}

// 获取当前 ISO 时间戳 (对应 datetime.datetime.now().isoformat())
std::string ExcelToolkitCore::GetIsoTimestampNow() {
    auto now = std::chrono::system_clock::now();
    auto in_time_t = std::chrono::system_clock::to_time_t(now);
    std::stringstream ss;
    ss << std::put_time(std::localtime(&in_time_t), "%Y-%m-%dT%H:%M:%S");
    return ss.str();
}

// 计算 Excel 标准 16-bit 密码 Hash (用于 sheetProtection password 属性)
uint16_t ExcelToolkitCore::ComputeExcelPasswordHash(const std::string& password) {
    if (password.empty()) return 0;
    uint16_t hash = 0;
    int len = (int)password.length();
    for (int i = len - 1; i >= 0; --i) {
        uint8_t c = (uint8_t)password[i];
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFE);
        hash ^= c;
    }
    hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFE);
    hash ^= (uint16_t)len;
    hash ^= 0xCE4B;
    return hash;
}

// 检查 soffice (LibreOffice) 是否可用
bool ExcelToolkitCore::CheckSofficeAvailable() {
#ifdef _WIN32
    int ret = system("where soffice >nul 2>nul");
    return ret == 0;
#else
    int ret = system("which soffice >/dev/null 2>&1");
    return ret == 0;
#endif
}

// 使用 soffice 进行格式互转 (xls <-> xlsx)
std::string ExcelToolkitCore::ConvertWithSoffice(const std::string& srcPath, const std::string& targetExt, const std::string& workDir) {
    std::string cmd = "soffice --headless --norestore --convert-to " + targetExt + " --outdir \"" + workDir + "\" \"" + srcPath + "\"";
    int ret = system(cmd.c_str());
    if (ret != 0) {
        throw std::runtime_error("soffice 转换命令执行失败");
    }
    fs::path p(srcPath);
    std::string stem = p.stem().u8string();
    fs::path outP = fs::u8path(workDir) / fs::u8path(stem + "." + targetExt);
    if (!fs::exists(outP)) {
        throw std::runtime_error("soffice 转换后未找到输出文件: " + outP.u8string());
    }
    return outP.u8string();
}

// ============================================================================
// ZIP 操作辅助类
// ============================================================================
struct ZipPackage {
    std::map<std::string, std::string> files;

    bool Load(const std::string& filePath, std::string& err) {
        mz_zip_archive zip;
        memset(&zip, 0, sizeof(zip));
        if (!mz_zip_reader_init_file(&zip, filePath.c_str(), 0)) {
            err = "无法以 ZIP 格式读取 XLSX 文件";
            return false;
        }

        uint32_t numFiles = mz_zip_reader_get_num_files(&zip);
        for (uint32_t i = 0; i < numFiles; ++i) {
            mz_zip_archive_file_stat stat;
            if (!mz_zip_reader_file_stat(&zip, i, &stat)) continue;
            if (stat.m_is_directory) continue;

            size_t size = 0;
            void* pData = mz_zip_reader_extract_to_heap(&zip, i, &size, 0);
            if (pData) {
                files[stat.m_filename] = std::string((char*)pData, size);
                mz_free(pData);
            }
        }
        mz_zip_reader_end(&zip);
        return true;
    }

    bool Save(const std::string& filePath, std::string& err) {
        std::string tmpPath = filePath + ".tmp_save";
        mz_zip_archive zip;
        memset(&zip, 0, sizeof(zip));
        if (!mz_zip_writer_init_file(&zip, tmpPath.c_str(), 0)) {
            err = "无法创建保存文件: " + tmpPath;
            return false;
        }

        for (const auto& kv : files) {
            mz_zip_writer_add_mem(&zip, kv.first.c_str(), kv.second.data(), kv.second.size(), MZ_DEFAULT_COMPRESSION);
        }

        if (!mz_zip_writer_finalize_archive(&zip)) {
            mz_zip_writer_end(&zip);
            err = "打包压缩包失败";
            return false;
        }
        mz_zip_writer_end(&zip);

        try {
            fs::rename(fs::u8path(tmpPath), fs::u8path(filePath));
        } catch (...) {
            std::ifstream src(tmpPath, std::ios::binary);
            std::ofstream dst(filePath, std::ios::binary);
            dst << src.rdbuf();
            src.close(); dst.close();
            fs::remove(fs::u8path(tmpPath));
        }
        return true;
    }
};

// ============================================================================
// 功能 1：添加暗水印 (Embed Watermark)
// ============================================================================
static bool EmbedWatermarkXlsx(const std::string& filePath, const std::string& watermarkText, const std::string& traceId, std::string& err) {
    ZipPackage pkg;
    if (!pkg.Load(filePath, err)) return false;

    if (pkg.files.find("xl/workbook.xml") == pkg.files.end()) {
        err = "缺少核心结构 xl/workbook.xml";
        return false;
    }

    std::string wbXml = pkg.files["xl/workbook.xml"];
    std::string relsXml = pkg.files["xl/_rels/workbook.xml.rels"];
    std::string ctXml = pkg.files["[Content_Types].xml"];

    std::string targetSheetFile = "worksheets/sheet_sysdata_cfg.xml";
    std::string targetSheetRelId = "rIdSysDataCfg";
    int targetSheetId = 9988;

    // 清理旧引用
    size_t oldSheetPos = wbXml.find("name=\"" + HIDDEN_SHEET_NAME_CONST + "\"");
    if (oldSheetPos != std::string::npos) {
        size_t tagStart = wbXml.rfind("<sheet", oldSheetPos);
        size_t tagEnd = wbXml.find("/>", oldSheetPos);
        if (tagEnd == std::string::npos) tagEnd = wbXml.find("</sheet>", oldSheetPos) + 8;
        else tagEnd += 2;
        if (tagStart != std::string::npos && tagEnd != std::string::npos) {
            wbXml.erase(tagStart, tagEnd - tagStart);
        }
    }

    // 1. 生成 ~sysdata_cfg 的 sheet XML 内容
    std::string nowIso = ExcelToolkitCore::GetIsoTimestampNow();
    std::stringstream ssSheet;
    ssSheet << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
            << "<worksheet xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/spreadsheets\">\n"
            << "  <dimension ref=\"A1:B3\"/>\n"
            << "  <sheetViews><sheetView tabSelected=\"0\" workbookViewId=\"0\"/></sheetViews>\n"
            << "  <sheetData>\n"
            << "    <row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>溯源编号</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>" << traceId << "</t></is></c></row>\n"
            << "    <row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>水印文字</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>" << watermarkText << "</t></is></c></row>\n"
            << "    <row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>生成时间</t></is></c><c r=\"B3\" t=\"inlineStr\"><is><t>" << nowIso << "</t></is></c></row>\n"
            << "  </sheetData>\n"
            << "</worksheet>";
    
    pkg.files["xl/" + targetSheetFile] = ssSheet.str();

    // 2. 更新 xl/workbook.xml 写入 state="veryHidden"
    std::string sheetTag = "<sheet name=\"" + HIDDEN_SHEET_NAME_CONST + "\" sheetId=\"" + std::to_string(targetSheetId) + "\" state=\"veryHidden\" r:id=\"" + targetSheetRelId + "\"/>";
    size_t sheetsEnd = wbXml.find("</sheets>");
    if (sheetsEnd != std::string::npos) {
        wbXml.insert(sheetsEnd, "  " + sheetTag + "\n");
    } else {
        wbXml = "<workbook><sheets>" + sheetTag + "</sheets></workbook>";
    }
    pkg.files["xl/workbook.xml"] = wbXml;

    // 3. 更新 xl/_rels/workbook.xml.rels
    if (relsXml.find("Id=\"" + targetSheetRelId + "\"") == std::string::npos) {
        std::string relTag = "<Relationship Id=\"" + targetSheetRelId + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" + targetSheetFile + "\"/>";
        size_t relsEnd = relsXml.find("</Relationships>");
        if (relsEnd != std::string::npos) {
            relsXml.insert(relsEnd, "  " + relTag + "\n");
            pkg.files["xl/_rels/workbook.xml.rels"] = relsXml;
        }
    }

    // 4. 更新 [Content_Types].xml
    if (ctXml.find("/xl/" + targetSheetFile) == std::string::npos) {
        std::string ctTag = "<Override PartName=\"/xl/" + targetSheetFile + "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
        size_t ctEnd = ctXml.find("</Types>");
        if (ctEnd != std::string::npos) {
            ctXml.insert(ctEnd, "  " + ctTag + "\n");
            pkg.files["[Content_Types].xml"] = ctXml;
        }
    }

    return pkg.Save(filePath, err);
}

bool ExcelToolkitCore::EmbedWatermark(
    const std::string& filePath,
    const std::string& watermarkText,
    const std::string& traceId,
    std::string& outError
) {
    fs::path p(filePath);
    std::string ext = p.extension().u8string();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

    if (ext == ".xlsx") {
        return EmbedWatermarkXlsx(filePath, watermarkText, traceId, outError);
    } else if (ext == ".xls") {
        try {
            fs::path tmpDir = fs::temp_directory_path() / fs::u8path("excel_tk_" + GenerateRandomHex8());
            fs::create_directories(tmpDir);

            std::string xlsxTmp = ConvertWithSoffice(filePath, "xlsx", tmpDir.u8string());
            if (!EmbedWatermarkXlsx(xlsxTmp, watermarkText, traceId, outError)) {
                fs::remove_all(tmpDir);
                return false;
            }
            std::string xlsTmp = ConvertWithSoffice(xlsxTmp, "xls", tmpDir.u8string());
            fs::copy_file(fs::u8path(xlsTmp), fs::u8path(filePath), fs::copy_options::overwrite_existing);
            fs::remove_all(tmpDir);
            return true;
        } catch (const std::exception& e) {
            outError = e.what();
            return false;
        }
    }

    outError = "仅支持 .xlsx 与 .xls 文件";
    return false;
}

// ============================================================================
// 功能 2：查看暗水印 (Inspect Watermark)
// ============================================================================
static bool InspectXlsx(const std::string& filePath, ExcelToolkitCore::InspectWatermarkInfo& outInfo, std::string& err) {
    ZipPackage pkg;
    if (!pkg.Load(filePath, err)) return false;

    if (pkg.files.find("xl/workbook.xml") == pkg.files.end()) {
        err = "缺少结构 xl/workbook.xml";
        return false;
    }

    std::string wbXml = pkg.files["xl/workbook.xml"];
    std::string relsXml = pkg.files["xl/_rels/workbook.xml.rels"];

    std::map<std::string, std::string> rIdToTarget;
    size_t rPos = 0;
    while ((rPos = relsXml.find("<Relationship ", rPos)) != std::string::npos) {
        size_t rEnd = relsXml.find("/>", rPos);
        if (rEnd == std::string::npos) break;
        std::string rTag = relsXml.substr(rPos, rEnd - rPos);

        size_t idPos = rTag.find("Id=\"");
        size_t targetPos = rTag.find("Target=\"");
        if (idPos != std::string::npos && targetPos != std::string::npos) {
            std::string id = rTag.substr(idPos + 4, rTag.find("\"", idPos + 4) - (idPos + 4));
            std::string target = rTag.substr(targetPos + 8, rTag.find("\"", targetPos + 8) - (targetPos + 8));
            rIdToTarget[id] = target;
        }
        rPos = rEnd;
    }

    size_t sPos = 0;
    std::string hiddenSheetXmlPath = "";

    while ((sPos = wbXml.find("<sheet ", sPos)) != std::string::npos) {
        size_t sEnd = wbXml.find("/>", sPos);
        if (sEnd == std::string::npos) sEnd = wbXml.find("</sheet>", sPos);
        if (sEnd == std::string::npos) break;

        std::string tag = wbXml.substr(sPos, sEnd - sPos);

        std::string name = "";
        size_t nPos = tag.find("name=\"");
        if (nPos != std::string::npos) {
            name = tag.substr(nPos + 6, tag.find("\"", nPos + 6) - (nPos + 6));
        }

        std::string state = "visible";
        size_t stPos = tag.find("state=\"");
        if (stPos != std::string::npos) {
            state = tag.substr(stPos + 7, tag.find("\"", stPos + 7) - (stPos + 7));
        }

        std::string rId = "";
        size_t ridPos = tag.find("r:id=\"");
        if (ridPos != std::string::npos) {
            rId = tag.substr(ridPos + 6, tag.find("\"", ridPos + 6) - (ridPos + 6));
        }

        ExcelToolkitCore::SheetMeta sm;
        sm.name = name;
        sm.state = state;
        outInfo.sheets.push_back(sm);

        if (name == HIDDEN_SHEET_NAME_CONST) {
            outInfo.hasHiddenSheet = true;
            if (rIdToTarget.find(rId) != rIdToTarget.end()) {
                std::string t = rIdToTarget[rId];
                if (t.rfind("worksheets/", 0) == 0) {
                    hiddenSheetXmlPath = "xl/" + t;
                } else if (t.rfind("xl/", 0) == 0) {
                    hiddenSheetXmlPath = t;
                } else {
                    hiddenSheetXmlPath = "xl/worksheets/" + t;
                }
            } else {
                hiddenSheetXmlPath = "xl/worksheets/sheet_sysdata_cfg.xml";
            }
        }

        sPos = sEnd;
    }

    if (outInfo.hasHiddenSheet && !hiddenSheetXmlPath.empty() && pkg.files.find(hiddenSheetXmlPath) != pkg.files.end()) {
        std::string hXml = pkg.files[hiddenSheetXmlPath];

        auto extractCellText = [&](const std::string& ref) -> std::string {
            size_t cPos = hXml.find("<c r=\"" + ref + "\"");
            if (cPos == std::string::npos) return "";
            size_t cEnd = hXml.find("</c>", cPos);
            if (cEnd == std::string::npos) return "";
            std::string cContent = hXml.substr(cPos, cEnd - cPos);
            size_t tPos = cContent.find("<t>");
            if (tPos == std::string::npos) tPos = cContent.find("<t ");
            if (tPos == std::string::npos) return "";
            size_t tValStart = cContent.find(">", tPos) + 1;
            size_t tValEnd = cContent.find("</t>", tValStart);
            if (tValEnd == std::string::npos) return "";
            return cContent.substr(tValStart, tValEnd - tValStart);
        };

        outInfo.traceId = extractCellText("B1");
        outInfo.watermarkText = extractCellText("B2");
        outInfo.createdAt = extractCellText("B3");
    }

    return true;
}

bool ExcelToolkitCore::InspectWatermarkFile(
    const std::string& filePath,
    InspectWatermarkInfo& outInfo,
    std::string& outError
) {
    fs::path p(filePath);
    std::string ext = p.extension().u8string();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

    if (ext == ".xlsx") {
        return InspectXlsx(filePath, outInfo, outError);
    } else if (ext == ".xls") {
        if (!CheckSofficeAvailable()) {
            outError = "此文件是 .xls 格式，查看需要依赖 LibreOffice 做格式转换，当前电脑未安装或未加入 PATH。";
            return false;
        }
        try {
            fs::path tmpDir = fs::temp_directory_path() / fs::u8path("excel_tk_" + GenerateRandomHex8());
            fs::create_directories(tmpDir);

            std::string xlsxTmp = ConvertWithSoffice(filePath, "xlsx", tmpDir.u8string());
            bool ok = InspectXlsx(xlsxTmp, outInfo, outError);
            fs::remove_all(tmpDir);
            return ok;
        } catch (const std::exception& e) {
            outError = e.what();
            return false;
        }
    }

    outError = "不支持的文件格式";
    return false;
}

// ============================================================================
// 功能 3：保护与解除保护 (Protect / Unprotect)
// ============================================================================
static bool ProtectXlsx(const std::string& filePath, const std::string& password, std::string& err) {
    ZipPackage pkg;
    if (!pkg.Load(filePath, err)) return false;

    uint16_t pwHash = ExcelToolkitCore::ComputeExcelPasswordHash(password);
    std::stringstream hexHash;
    hexHash << std::hex << std::uppercase << std::setw(4) << std::setfill('0') << pwHash;
    std::string pwHashHex = hexHash.str();

    std::string protTag = "<sheetProtection password=\"" + pwHashHex + "\" sheet=\"1\" objects=\"1\" scenarios=\"1\" formatCells=\"0\" formatColumns=\"0\" formatRows=\"0\" insertColumns=\"0\" insertRows=\"0\" insertHyperlinks=\"0\" deleteColumns=\"0\" deleteRows=\"0\" selectLockedCells=\"1\" sort=\"0\" autoFilter=\"0\" pivotTables=\"0\" selectUnlockedCells=\"1\"/>";

    for (auto& kv : pkg.files) {
        if (kv.first.rfind("xl/worksheets/sheet", 0) == 0 && kv.first.find(".xml") != std::string::npos) {
            std::string& sXml = kv.second;
            size_t oldProt = sXml.find("<sheetProtection ");
            if (oldProt != std::string::npos) {
                size_t pEnd = sXml.find("/>", oldProt);
                if (pEnd != std::string::npos) sXml.erase(oldProt, pEnd + 2 - oldProt);
            }
            size_t insertPos = sXml.find("</worksheet>");
            if (insertPos != std::string::npos) {
                sXml.insert(insertPos, "  " + protTag + "\n");
            }
        }
    }

    if (pkg.files.find("xl/workbook.xml") != pkg.files.end()) {
        std::string& wbXml = pkg.files["xl/workbook.xml"];
        size_t oldWbProt = wbXml.find("<workbookProtection ");
        if (oldWbProt != std::string::npos) {
            size_t pEnd = wbXml.find("/>", oldWbProt);
            if (pEnd != std::string::npos) wbXml.erase(oldWbProt, pEnd + 2 - oldWbProt);
        }

        std::string wbProtTag = "<workbookProtection workbookPassword=\"" + pwHashHex + "\" lockStructure=\"1\" lockWindows=\"0\"/>";
        size_t insertPos = wbXml.find("</workbook>");
        if (insertPos != std::string::npos) {
            wbXml.insert(insertPos, "  " + wbProtTag + "\n");
        }
    }

    return pkg.Save(filePath, err);
}

static bool UnprotectXlsx(const std::string& filePath, std::string& err) {
    ZipPackage pkg;
    if (!pkg.Load(filePath, err)) return false;

    for (auto& kv : pkg.files) {
        if (kv.first.rfind("xl/worksheets/sheet", 0) == 0 && kv.first.find(".xml") != std::string::npos) {
            std::string& sXml = kv.second;
            size_t oldProt = sXml.find("<sheetProtection");
            while ((oldProt = sXml.find("<sheetProtection", oldProt)) != std::string::npos) {
                size_t pEnd = sXml.find("/>", oldProt);
                if (pEnd != std::string::npos) {
                    sXml.erase(oldProt, pEnd + 2 - oldProt);
                } else {
                    break;
                }
            }
        }
    }

    if (pkg.files.find("xl/workbook.xml") != pkg.files.end()) {
        std::string& wbXml = pkg.files["xl/workbook.xml"];
        size_t oldWbProt = wbXml.find("<workbookProtection");
        while ((oldWbProt = wbXml.find("<workbookProtection", oldWbProt)) != std::string::npos) {
            size_t pEnd = wbXml.find("/>", oldWbProt);
            if (pEnd != std::string::npos) {
                wbXml.erase(oldWbProt, pEnd + 2 - oldWbProt);
            } else {
                break;
            }
        }
    }

    return pkg.Save(filePath, err);
}

bool ExcelToolkitCore::ProtectWorkbook(
    const std::string& filePath,
    const std::string& password,
    std::string& outError
) {
    fs::path p(filePath);
    std::string ext = p.extension().u8string();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

    if (ext == ".xlsx") {
        return ProtectXlsx(filePath, password, outError);
    } else if (ext == ".xls") {
        try {
            fs::path tmpDir = fs::temp_directory_path() / fs::u8path("excel_tk_" + GenerateRandomHex8());
            fs::create_directories(tmpDir);

            std::string xlsxTmp = ConvertWithSoffice(filePath, "xlsx", tmpDir.u8string());
            if (!ProtectXlsx(xlsxTmp, password, outError)) {
                fs::remove_all(tmpDir);
                return false;
            }
            std::string xlsTmp = ConvertWithSoffice(xlsxTmp, "xls", tmpDir.u8string());
            fs::copy_file(fs::u8path(xlsTmp), fs::u8path(filePath), fs::copy_options::overwrite_existing);
            fs::remove_all(tmpDir);
            return true;
        } catch (const std::exception& e) {
            outError = e.what();
            return false;
        }
    }

    outError = "不支持的文件格式";
    return false;
}

bool ExcelToolkitCore::UnprotectWorkbook(
    const std::string& filePath,
    std::string& outError
) {
    fs::path p(filePath);
    std::string ext = p.extension().u8string();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);

    if (ext == ".xlsx") {
        return UnprotectXlsx(filePath, outError);
    } else if (ext == ".xls") {
        try {
            fs::path tmpDir = fs::temp_directory_path() / fs::u8path("excel_tk_" + GenerateRandomHex8());
            fs::create_directories(tmpDir);

            std::string xlsxTmp = ConvertWithSoffice(filePath, "xlsx", tmpDir.u8string());
            if (!UnprotectXlsx(xlsxTmp, outError)) {
                fs::remove_all(tmpDir);
                return false;
            }
            std::string xlsTmp = ConvertWithSoffice(xlsxTmp, "xls", tmpDir.u8string());
            fs::copy_file(fs::u8path(xlsTmp), fs::u8path(filePath), fs::copy_options::overwrite_existing);
            fs::remove_all(tmpDir);
            return true;
        } catch (const std::exception& e) {
            outError = e.what();
            return false;
        }
    }

    outError = "不支持的文件格式";
    return false;
}
