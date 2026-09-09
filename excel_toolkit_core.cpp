/**
 * ============================================================================
 * excel_toolkit_core.cpp - OpenXML / ZIP 内存安全深度实现
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
#include <cstring>

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#endif

namespace fs = std::filesystem;

static bool ReadEntireFile(const std::string& u8path, std::vector<uint8_t>& outBuffer, std::string& err) {
    try {
#ifdef _WIN32
        int wlen = MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, NULL, 0);
        if (wlen <= 0) { err = "路径编码解析失败"; return false; }
        std::wstring wpath(wlen - 1, 0);
        MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, &wpath[0], wlen);
        FILE* fp = _wfopen(wpath.c_str(), L"rb");
#else
        FILE* fp = fopen(u8path.c_str(), "rb");
#endif
        if (!fp) {
            err = "无法打开文件，请确认文件存在且未被 Excel 占用锁定。";
            return false;
        }
        fseek(fp, 0, SEEK_END);
        long sz = ftell(fp);
        fseek(fp, 0, SEEK_SET);
        if (sz < 0) {
            fclose(fp);
            err = "获取文件大小失败";
            return false;
        }
        outBuffer.resize((size_t)sz);
        if (sz > 0) {
            size_t readBytes = fread(outBuffer.data(), 1, (size_t)sz, fp);
            if (readBytes != (size_t)sz) {
                fclose(fp);
                err = "读取文件内容不完整";
                return false;
            }
        }
        fclose(fp);
        return true;
    } catch (const std::exception& e) {
        err = std::string("读取文件异常: ") + e.what();
        return false;
    } catch (...) {
        err = "读取文件发生未知异常";
        return false;
    }
}

static bool WriteEntireFile(const std::string& u8path, const void* data, size_t size, std::string& err) {
    try {
#ifdef _WIN32
        int wlen = MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, NULL, 0);
        if (wlen <= 0) { err = "路径编码解析失败"; return false; }
        std::wstring wpath(wlen - 1, 0);
        MultiByteToWideChar(CP_UTF8, 0, u8path.c_str(), -1, &wpath[0], wlen);
        FILE* fp = _wfopen(wpath.c_str(), L"wb");
#else
        FILE* fp = fopen(u8path.c_str(), "wb");
#endif
        if (!fp) {
            err = "无法写入目标文件，请确认目标文件未被其他程序打开且拥有写入权限。";
            return false;
        }
        if (size > 0 && data != nullptr) {
            size_t written = fwrite(data, 1, size, fp);
            if (written != size) {
                fclose(fp);
                err = "写入文件数据不完整";
                return false;
            }
        }
        fclose(fp);
        return true;
    } catch (const std::exception& e) {
        err = std::string("写入文件异常: ") + e.what();
        return false;
    } catch (...) {
        err = "写入文件发生未知异常";
        return false;
    }
}

std::string ExcelToolkitCore::GenerateRandomHex8() {
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<uint32_t> dis(0, 0xFFFFFFFF);
    uint32_t val = dis(gen);
    std::stringstream ss;
    ss << std::hex << std::setw(8) << std::setfill('0') << val;
    return ss.str();
}

std::string ExcelToolkitCore::GetIsoTimestampNow() {
    auto now = std::chrono::system_clock::now();
    auto in_time_t = std::chrono::system_clock::to_time_t(now);
    // 强制转换为中国标准时间 CST (UTC+8)
    std::time_t cst_time = in_time_t + 8 * 3600;
    std::tm* gmt = std::gmtime(&cst_time);
    char buf[64] = { 0 };
    if (gmt) {
        std::strftime(buf, sizeof(buf), "%Y-%m-%d %H:%M:%S (CST)", gmt);
        return std::string(buf);
    }
    return "2026-09-09 15:00:00 (CST)";
}

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

bool ExcelToolkitCore::CheckSofficeAvailable() {
#ifdef _WIN32
    int ret = system("where soffice >nul 2>nul");
    return ret == 0;
#else
    int ret = system("which soffice >/dev/null 2>&1");
    return ret == 0;
#endif
}

std::string ExcelToolkitCore::ConvertWithSoffice(const std::string& srcPath, const std::string& targetExt, const std::string& workDir) {
    std::string cmd = "soffice --headless --norestore --convert-to " + targetExt + " --outdir \"" + workDir + "\" \"" + srcPath + "\"";
    int ret = system(cmd.c_str());
    if (ret != 0) {
        throw std::runtime_error("soffice 转换命令执行失败");
    }
    fs::path p = fs::u8path(srcPath);
    std::string stem = p.stem().u8string();
    fs::path outP = fs::u8path(workDir) / fs::u8path(stem + "." + targetExt);
    if (!fs::exists(outP)) {
        throw std::runtime_error("soffice 转换后未找到输出文件: " + outP.u8string());
    }
    return outP.u8string();
}

struct MemoryZipPackage {
    std::map<std::string, std::string> files;

    bool LoadFromBuffer(const std::vector<uint8_t>& buffer, std::string& err) {
        if (buffer.size() < 22) {
            err = "文件过小，非有效 Excel 压缩容器。";
            return false;
        }

        mz_zip_archive zip;
        memset(&zip, 0, sizeof(zip));
        if (!mz_zip_reader_init_mem(&zip, buffer.data(), buffer.size(), 0)) {
            err = "无法以 ZIP 容器解析该 XLSX 文件（可能损坏或非标准 xlsx 格式）。";
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

    bool SaveToFile(const std::string& outputPath, std::string& err) {
        mz_zip_archive zip;
        memset(&zip, 0, sizeof(zip));
        if (!mz_zip_writer_init_heap(&zip, 0, 0)) {
            err = "初始化内存 ZIP 压缩器失败";
            return false;
        }

        for (const auto& kv : files) {
            mz_zip_writer_add_mem(&zip, kv.first.c_str(), kv.second.data(), kv.second.size(), MZ_DEFAULT_COMPRESSION);
        }

        void* pBuf = nullptr;
        size_t bufSize = 0;
        if (!mz_zip_writer_finalize_heap_archive(&zip, &pBuf, &bufSize)) {
            mz_zip_writer_end(&zip);
            err = "打包 XLSX 数据失败";
            return false;
        }

        bool writeOk = WriteEntireFile(outputPath, pBuf, bufSize, err);
        mz_zip_writer_end(&zip);
        return writeOk;
    }
};

bool ExcelToolkitCore::ProcessExcelFile(
    const std::string& inputPath,
    const std::string& outputPath,
    const ProcessOptions& options,
    std::string& outError
) {
    try {
        std::vector<uint8_t> inBuffer;
        if (!ReadEntireFile(inputPath, inBuffer, outError)) {
            return false;
        }

        MemoryZipPackage pkg;
        if (!pkg.LoadFromBuffer(inBuffer, outError)) {
            return false;
        }

        if (pkg.files.find("xl/workbook.xml") == pkg.files.end()) {
            outError = "未找到核心文件 xl/workbook.xml，该文件可能不是标准的 XLSX 文件。";
            return false;
        }

        std::string& wbXml = pkg.files["xl/workbook.xml"];
        std::string& relsXml = pkg.files["xl/_rels/workbook.xml.rels"];
        std::string& ctXml = pkg.files["[Content_Types].xml"];

        if (options.unprotectAll) {
            size_t wpPos = wbXml.find("<workbookProtection");
            if (wpPos != std::string::npos) {
                size_t wpEnd = wbXml.find("/>", wpPos);
                if (wpEnd != std::string::npos) {
                    wbXml.erase(wpPos, wpEnd + 2 - wpPos);
                }
            }

            for (auto& kv : pkg.files) {
                if (kv.first.rfind("xl/worksheets/sheet", 0) == 0 && kv.first.find(".xml") != std::string::npos) {
                    std::string& sXml = kv.second;
                    size_t spPos = sXml.find("<sheetProtection");
                    if (spPos != std::string::npos) {
                        size_t spEnd = sXml.find("/>", spPos);
                        if (spEnd != std::string::npos) {
                            sXml.erase(spPos, spEnd + 2 - spPos);
                        }
                    }
                }
            }
        } else {
            if (options.lockWorkbookStructure) {
                size_t wpPos = wbXml.find("<workbookProtection");
                if (wpPos != std::string::npos) {
                    size_t wpEnd = wbXml.find("/>", wpPos);
                    if (wpEnd != std::string::npos) {
                        wbXml.erase(wpPos, wpEnd + 2 - wpPos);
                    }
                }
                std::string wbProt = "<workbookProtection lockStructure=\"1\" lockWindows=\"0\"/>";
                size_t insertPos = wbXml.find("</workbook>");
                if (insertPos != std::string::npos) {
                    wbXml.insert(insertPos, wbProt);
                }
            }

            if (options.protectSheetContent) {
                uint16_t pwHash = ComputeExcelPasswordHash(options.password);
                std::stringstream hexHash;
                hexHash << std::hex << std::uppercase << std::setw(4) << std::setfill('0') << pwHash;
                std::string pwHashHex = hexHash.str();

                std::string protTag = "<sheetProtection password=\"" + pwHashHex + "\" sheet=\"1\" objects=\"1\" scenarios=\"1\" formatCells=\"0\" formatColumns=\"0\" formatRows=\"0\" insertColumns=\"0\" insertRows=\"0\" insertHyperlinks=\"0\" deleteColumns=\"0\" deleteRows=\"0\" selectLockedCells=\"1\" sort=\"0\" autoFilter=\"0\" pivotTables=\"0\" selectUnlockedCells=\"1\"/>";

                for (auto& kv : pkg.files) {
                    if (kv.first.rfind("xl/worksheets/sheet", 0) == 0 && kv.first.find(".xml") != std::string::npos) {
                        std::string& sXml = kv.second;
                        size_t spPos = sXml.find("<sheetProtection");
                        if (spPos != std::string::npos) {
                            size_t spEnd = sXml.find("/>", spPos);
                            if (spEnd != std::string::npos) {
                                sXml.erase(spPos, spEnd + 2 - spPos);
                            }
                        }
                        size_t insPos = sXml.find("</worksheet>");
                        if (insPos != std::string::npos) {
                            sXml.insert(insPos, protTag);
                        }
                    }
                }
            }
        }

        if (options.addWatermark) {
            std::string effectiveSheetName = options.sheetName.empty() ? HIDDEN_SHEET_NAME_CONST : options.sheetName;
            std::string effectiveWmText = options.watermarkText.empty() ? "绝密外发审计凭证" : options.watermarkText;
            std::string effectiveTraceId = options.traceId.empty() ? ("SEC-" + GenerateRandomHex8()) : options.traceId;
            std::string nowIso = GetIsoTimestampNow();

            std::string targetSheetFile = "worksheets/sheet_sysdata_cfg.xml";
            std::string targetSheetRelId = "rIdSysDataCfg";
            int targetSheetId = 9988;

            size_t oldSheetPos = wbXml.find("name=\"" + effectiveSheetName + "\"");
            if (oldSheetPos != std::string::npos) {
                size_t tagStart = wbXml.rfind("<sheet", oldSheetPos);
                size_t tagEnd = wbXml.find("/>", oldSheetPos);
                if (tagEnd == std::string::npos) tagEnd = wbXml.find("</sheet>", oldSheetPos) + 8;
                else tagEnd += 2;
                if (tagStart != std::string::npos && tagEnd != std::string::npos) {
                    wbXml.erase(tagStart, tagEnd - tagStart);
                }
            }

            std::stringstream ssSheet;
            ssSheet << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                    << "<worksheet xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/spreadsheets\">\n"
                    << "  <dimension ref=\"A1:B3\"/>\n"
                    << "  <sheetViews><sheetView tabSelected=\"0\" workbookViewId=\"0\"/></sheetViews>\n"
                    << "  <sheetData>\n"
                    << "    <row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>溯源编号</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>" << effectiveTraceId << "</t></is></c></row>\n"
                    << "    <row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>水印文字</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>" << effectiveWmText << "</t></is></c></row>\n"
                    << "    <row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>生成时间</t></is></c><c r=\"B3\" t=\"inlineStr\"><is><t>" << nowIso << "</t></is></c></row>\n"
                    << "  </sheetData>\n"
                    << "</worksheet>";

            pkg.files["xl/" + targetSheetFile] = ssSheet.str();

            std::string sheetTag = "<sheet name=\"" + effectiveSheetName + "\" sheetId=\"" + std::to_string(targetSheetId) + "\" state=\"veryHidden\" r:id=\"" + targetSheetRelId + "\"/>";
            size_t sheetsEnd = wbXml.find("</sheets>");
            if (sheetsEnd != std::string::npos) {
                wbXml.insert(sheetsEnd, "  " + sheetTag + "\n");
            } else {
                wbXml = "<workbook><sheets>" + sheetTag + "</sheets></workbook>";
            }

            if (relsXml.find("Id=\"" + targetSheetRelId + "\"") == std::string::npos) {
                std::string relTag = "<Relationship Id=\"" + targetSheetRelId + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" + targetSheetFile + "\"/>";
                size_t relsEnd = relsXml.find("</Relationships>");
                if (relsEnd != std::string::npos) {
                    relsXml.insert(relsEnd, "  " + relTag + "\n");
                }
            }

            if (ctXml.find("/xl/" + targetSheetFile) == std::string::npos) {
                std::string ctTag = "<Override PartName=\"/xl/" + targetSheetFile + "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
                size_t ctEnd = ctXml.find("</Types>");
                if (ctEnd != std::string::npos) {
                    ctXml.insert(ctEnd, "  " + ctTag + "\n");
                }
            }
        }

        std::string finalOutput = outputPath.empty() ? inputPath : outputPath;
        return pkg.SaveToFile(finalOutput, outError);
    } catch (const std::exception& e) {
        outError = std::string("处理异常: ") + e.what();
        return false;
    } catch (...) {
        outError = "处理发生未知内部异常";
        return false;
    }
}

bool ExcelToolkitCore::InspectWatermarkFile(
    const std::string& filePath,
    InspectWatermarkInfo& outInfo,
    std::string& outError
) {
    try {
        std::vector<uint8_t> inBuffer;
        if (!ReadEntireFile(filePath, inBuffer, outError)) {
            return false;
        }

        MemoryZipPackage pkg;
        if (!pkg.LoadFromBuffer(inBuffer, outError)) {
            return false;
        }

        if (pkg.files.find("xl/workbook.xml") == pkg.files.end()) {
            outError = "缺少核心结构 xl/workbook.xml";
            return false;
        }

        const std::string& wbXml = pkg.files["xl/workbook.xml"];
        std::string relsXml = "";
        if (pkg.files.find("xl/_rels/workbook.xml.rels") != pkg.files.end()) {
            relsXml = pkg.files["xl/_rels/workbook.xml.rels"];
        }

        if (wbXml.find("lockStructure=\"1\"") != std::string::npos || wbXml.find("lockStructure=\"true\"") != std::string::npos) {
            outInfo.isWorkbookLocked = true;
        }

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

        // 解析共享字符池 (xl/sharedStrings.xml)
        std::vector<std::string> sharedStrings;
        if (pkg.files.find("xl/sharedStrings.xml") != pkg.files.end()) {
            const std::string& ssXml = pkg.files["xl/sharedStrings.xml"];
            size_t siPos = 0;
            while ((siPos = ssXml.find("<si", siPos)) != std::string::npos) {
                size_t siEnd = ssXml.find("</si>", siPos);
                if (siEnd == std::string::npos) break;
                std::string siChunk = ssXml.substr(siPos, siEnd - siPos);
                std::string tCombined = "";
                size_t tp = 0;
                while ((tp = siChunk.find("<t", tp)) != std::string::npos) {
                    size_t valStart = siChunk.find(">", tp);
                    if (valStart == std::string::npos) break;
                    valStart++;
                    size_t valEnd = siChunk.find("</t>", valStart);
                    if (valEnd != std::string::npos) {
                        tCombined += siChunk.substr(valStart, valEnd - valStart);
                        tp = valEnd + 4;
                    } else {
                        break;
                    }
                }
                sharedStrings.push_back(tCombined);
                siPos = siEnd + 5;
            }
        }

        auto findSecPattern = [](const std::string& str) -> std::string {
            size_t pos = str.find("SEC-");
            if (pos == std::string::npos) return "";
            size_t endPos = pos + 4;
            while (endPos < str.size() && (isalnum((unsigned char)str[endPos]) || str[endPos] == '-' || str[endPos] == '_')) {
                endPos++;
            }
            return str.substr(pos, endPos - pos);
        };

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

            SheetMeta sm;
            sm.name = name;
            sm.state = state;
            outInfo.sheets.push_back(sm);

            if (name == HIDDEN_SHEET_NAME_CONST || state == "veryHidden") {
                outInfo.hasHiddenSheet = true;
                if (rIdToTarget.find(rId) != rIdToTarget.end()) {
                    std::string t = rIdToTarget[rId];
                    while (!t.empty() && (t[0] == '/' || t[0] == '\\')) t.erase(0, 1);
                    if (t.rfind("./", 0) == 0) t.erase(0, 2);

                    if (t.rfind("xl/", 0) == 0) {
                        hiddenSheetXmlPath = t;
                    } else if (t.rfind("worksheets/", 0) == 0) {
                        hiddenSheetXmlPath = "xl/" + t;
                    } else {
                        hiddenSheetXmlPath = "xl/worksheets/" + t;
                    }
                } else {
                    hiddenSheetXmlPath = "xl/worksheets/sheet_sysdata_cfg.xml";
                }
            }

            sPos = sEnd;
        }

        auto extractCellText = [&](const std::string& ref, const std::string& xml) -> std::string {
            size_t cPos = xml.find("<c r=\"" + ref + "\"");
            if (cPos == std::string::npos) return "";
            size_t cEnd = xml.find("</c>", cPos);
            if (cEnd == std::string::npos) cEnd = xml.find("/>", cPos);
            if (cEnd == std::string::npos) return "";
            std::string cContent = xml.substr(cPos, cEnd - cPos + 4);

            if (cContent.find("t=\"s\"") != std::string::npos) {
                size_t vPos = cContent.find("<v>");
                if (vPos != std::string::npos) {
                    size_t vEnd = cContent.find("</v>", vPos);
                    if (vEnd != std::string::npos) {
                        std::string idxStr = cContent.substr(vPos + 3, vEnd - (vPos + 3));
                        try {
                            size_t idx = std::stoul(idxStr);
                            if (idx < sharedStrings.size()) {
                                return sharedStrings[idx];
                            }
                        } catch (...) {}
                    }
                }
            }

            size_t tPos = cContent.find("<t>");
            if (tPos == std::string::npos) tPos = cContent.find("<t ");
            if (tPos != std::string::npos) {
                size_t tValStart = cContent.find(">", tPos) + 1;
                size_t tValEnd = cContent.find("</t>", tValStart);
                if (tValEnd != std::string::npos) {
                    return cContent.substr(tValStart, tValEnd - tValStart);
                }
            }

            size_t vPos = cContent.find("<v>");
            if (vPos != std::string::npos) {
                size_t vEnd = cContent.find("</v>", vPos);
                if (vEnd != std::string::npos) {
                    return cContent.substr(vPos + 3, vEnd - (vPos + 3));
                }
            }
            return "";
        };

        if (outInfo.hasHiddenSheet) {
            if (hiddenSheetXmlPath.empty() || pkg.files.find(hiddenSheetXmlPath) == pkg.files.end()) {
                for (const auto& kv : pkg.files) {
                    if (kv.first.find("sheet_sysdata_cfg") != std::string::npos) {
                        hiddenSheetXmlPath = kv.first;
                        break;
                    }
                }
            }

            if (!hiddenSheetXmlPath.empty() && pkg.files.find(hiddenSheetXmlPath) != pkg.files.end()) {
                const std::string& hXml = pkg.files[hiddenSheetXmlPath];
                outInfo.traceId = extractCellText("B1", hXml);
                outInfo.watermarkText = extractCellText("B2", hXml);
                outInfo.createdAt = extractCellText("B3", hXml);

                if (outInfo.traceId.empty()) {
                    outInfo.traceId = findSecPattern(hXml);
                }
            }
        }

        // 全盘检索：若 traceId 仍未抓取到，从 sharedStrings 及所有工作表 XML 中检索 SEC- 溯源码
        if (outInfo.traceId.empty()) {
            for (const auto& kv : pkg.files) {
                if (kv.first.find("xl/worksheets/") != std::string::npos ||
                    kv.first == "xl/sharedStrings.xml" ||
                    kv.first == "xl/workbook.xml") {
                    std::string sec = findSecPattern(kv.second);
                    if (!sec.empty()) {
                        outInfo.traceId = sec;
                        outInfo.hasHiddenSheet = true;
                        break;
                    }
                }
            }
        }

        return true;
    } catch (const std::exception& e) {
        outError = std::string("解析异常: ") + e.what();
        return false;
    } catch (...) {
        outError = "解析发生未知内部异常";
        return false;
    }
}

bool ExcelToolkitCore::EmbedWatermark(
    const std::string& filePath,
    const std::string& watermarkText,
    const std::string& traceId,
    std::string& outError
) {
    ProcessOptions opt;
    opt.addWatermark = true;
    opt.watermarkText = watermarkText;
    opt.traceId = traceId;
    opt.lockWorkbookStructure = true;
    return ProcessExcelFile(filePath, filePath, opt, outError);
}

bool ExcelToolkitCore::ProtectWorkbook(
    const std::string& filePath,
    const std::string& password,
    std::string& outError
) {
    ProcessOptions opt;
    opt.addWatermark = false;
    opt.lockWorkbookStructure = true;
    opt.protectSheetContent = true;
    opt.password = password;
    return ProcessExcelFile(filePath, filePath, opt, outError);
}

bool ExcelToolkitCore::UnprotectWorkbook(
    const std::string& filePath,
    std::string& outError
) {
    ProcessOptions opt;
    opt.addWatermark = false;
    opt.unprotectAll = true;
    return ProcessExcelFile(filePath, filePath, opt, outError);
}
